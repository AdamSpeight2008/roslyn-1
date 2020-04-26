' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Function BindQueryOperatorCall(
                                                node As SyntaxNode,
                                                source As BoundExpression,
                                                operatorName As String,
                                                arguments As ImmutableArray(Of BoundExpression),
                                                operatorNameLocation As TextSpan,
                                                diagnostics As DiagnosticBag
                                              ) As BoundExpression
            Return BindQueryOperatorCall(node,
                                         source,
                                         operatorName,
                                         LookupQueryOperator(node, source, operatorName, Nothing, diagnostics),
                                         arguments,
                                         operatorNameLocation,
                                         diagnostics)
        End Function

        Private Function BindQueryOperatorCall(
                                                node As SyntaxNode,
                                                source As BoundExpression,
                                                operatorName As String,
                                                typeArgumentsOpt As BoundTypeArguments,
                                                arguments As ImmutableArray(Of BoundExpression),
                                                operatorNameLocation As TextSpan,
                                                diagnostics As DiagnosticBag
                                              ) As BoundExpression
            Return BindQueryOperatorCall(node,
                                         source,
                                         operatorName,
                                         LookupQueryOperator(node, source, operatorName, typeArgumentsOpt, diagnostics),
                                         arguments,
                                         operatorNameLocation,
                                         diagnostics)
        End Function

        ''' <summary> <paramref name="methodGroup"/> can be Nothing if lookup didn't find anything. </summary>
        Private Function BindQueryOperatorCall(
                                                node As SyntaxNode,
                                                source As BoundExpression,
                                                operatorName As String,
                                                methodGroup As BoundMethodGroup,
                                                arguments As ImmutableArray(Of BoundExpression),
                                                operatorNameLocation As TextSpan,
                                                diagnostics As DiagnosticBag
                                              ) As BoundExpression
            Debug.Assert(source.IsValue)
            Debug.Assert(methodGroup Is Nothing OrElse
                         (methodGroup.ReceiverOpt Is source AndAlso
                            (methodGroup.ResultKind = LookupResultKind.Good OrElse methodGroup.ResultKind = LookupResultKind.Inaccessible)))

            Dim boundCall As BoundExpression = Nothing

            If methodGroup IsNot Nothing Then
                Query_BindMethodGroup(node, methodGroup, arguments, operatorNameLocation, diagnostics, boundCall)
            End If

            If boundCall Is Nothing Then
                boundCall = Query_BindBoundCall(node, source, methodGroup, arguments)
            End If

            If boundCall.HasErrors AndAlso Not source.HasErrors Then
                ReportDiagnostic(diagnostics, Location.Create(node.SyntaxTree, operatorNameLocation), ERRID.ERR_QueryOperatorNotFound, operatorName)
            End If

            boundCall.SetWasCompilerGenerated()

            Return boundCall
        End Function

        Private Shared Function Query_BindBoundCall(
                                                     node As SyntaxNode,
                                                     source As BoundExpression,
                                                     methodGroup As BoundMethodGroup,
                                                     arguments As ImmutableArray(Of BoundExpression)
                                                   ) As BoundExpression
            Dim boundCall As BoundExpression
            Dim childBoundNodes As ImmutableArray(Of BoundExpression)

            If arguments.IsEmpty Then
                childBoundNodes = ImmutableArray.Create(If(methodGroup, source))
            Else
                Dim builder = ArrayBuilder(Of BoundExpression).GetInstance()
                With builder
                    .Add(If(methodGroup, source))
                    .AddRange(arguments)
                    childBoundNodes = .ToImmutableAndFree()
                End With
            End If

            If methodGroup Is Nothing Then
                boundCall = BadExpression(node, childBoundNodes, ErrorTypeSymbol.UnknownResultType)
            Else
                Dim symbols = ArrayBuilder(Of Symbol).GetInstance()
                methodGroup.GetExpressionSymbols(symbols)

                Dim resultKind = LookupResultKind.OverloadResolutionFailure
                If methodGroup.ResultKind < resultKind Then resultKind = methodGroup.ResultKind

                boundCall = New BoundBadExpression(node,
                                                   resultKind,
                                                   symbols.ToImmutableAndFree(),
                                                   childBoundNodes,
                                                   ErrorTypeSymbol.UnknownResultType,
                                                   hasErrors:=True)
            End If

            Return boundCall
        End Function

        Private Sub Query_BindMethodGroup(
                                           node As SyntaxNode,
                                           methodGroup As BoundMethodGroup,
                                           arguments As ImmutableArray(Of BoundExpression),
                                           operatorNameLocation As TextSpan,
                                     ByRef diagnostics As DiagnosticBag,
                                     ByRef boundCall As BoundExpression
                                         )

            Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
            Dim results = OverloadResolution.QueryOperatorInvocationOverloadResolution(methodGroup,
                                                                                       arguments,
                                                                                       Me,
                                                                                       useSiteDiagnostics)

            If diagnostics.Add(node, useSiteDiagnostics) Then
                If methodGroup.ResultKind <> LookupResultKind.Inaccessible Then
                    ' Suppress additional diagnostics
                    diagnostics = New DiagnosticBag()
                End If
            End If

            If Not results.BestResult.HasValue Then
                ' Create and report the diagnostic.
                If results.Candidates.Length = 0 Then
                    results = OverloadResolution.QueryOperatorInvocationOverloadResolution(methodGroup,
                                                                                           arguments,
                                                                                           Me,
                                                                                           includeEliminatedCandidates:=True,
                                                                                           useSiteDiagnostics:=useSiteDiagnostics)
                End If

                If results.Candidates.Length > 0 Then
                    boundCall = ReportOverloadResolutionFailureAndProduceBoundNode(node,
                                                                                   methodGroup,
                                                                                   arguments,
                                                                                   Nothing,
                                                                                   results,
                                                                                   diagnostics,
                                                                                   callerInfoOpt:=Nothing,
                                                                                   queryMode:=True,
                                                                                   diagnosticLocationOpt:=Location.Create(node.SyntaxTree, operatorNameLocation))
                End If
            Else
                boundCall = CreateBoundCallOrPropertyAccess(node,
                                                            node,
                                                            TypeCharacter.None,
                                                            methodGroup,
                                                            arguments,
                                                            results.BestResult.Value,
                                                            results.AsyncLambdaSubToFunctionMismatch,
                                                            diagnostics)

                ' May need to update return type for LambdaSymbols associated with query lambdas.
                For i As Integer = 0 To arguments.Length - 1
                    Dim arg As BoundExpression = arguments(i)

                    If arg.Kind = BoundKind.QueryLambda Then
                        Dim queryLambda = DirectCast(arg, BoundQueryLambda)

                        If queryLambda.LambdaSymbol.ReturnType Is LambdaSymbol.ReturnTypePendingDelegate Then
                            Dim delegateReturnType As TypeSymbol = DirectCast(boundCall, BoundCall).Method.Parameters(i).Type.DelegateOrExpressionDelegate(Me).DelegateInvokeMethod.ReturnType
                            queryLambda.LambdaSymbol.SetQueryLambdaReturnType(delegateReturnType)
                        End If
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Return method group or Nothing in case nothing was found.
        ''' Note, returned group might have ResultKind = "Inaccessible".
        ''' </summary>
        Private Function LookupQueryOperator(
                                              node As SyntaxNode,
                                              source As BoundExpression,
                                              operatorName As String,
                                              typeArgumentsOpt As BoundTypeArguments,
                                              diagnostics As DiagnosticBag
                                            ) As BoundMethodGroup

            Dim lookupResult As LookupResult = LookupResult.GetInstance()
            Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
            LookupMember(lookupResult, source.Type, operatorName, 0, QueryOperatorLookupOptions, useSiteDiagnostics)

            Dim methodGroup As BoundMethodGroup = Nothing

            ' NOTE: Lookup may return Kind = LookupResultKind.Inaccessible or LookupResultKind.MustBeInstance;
            '
            '       It looks we intentionally pass LookupResultKind.Inaccessible to CreateBoundMethodGroup(...) 
            '       causing BC30390 to be generated instead of BC36594 reported by Dev11 (more accurate message?)
            '
            '       As CreateBoundMethodGroup(...) only expects Kind = LookupResultKind.Good or 
            '       LookupResultKind.Inaccessible in all other cases we just skip calling this method
            '       so that BC36594 is generated which what seems to what Dev11 does.
            If Not lookupResult.IsClear AndAlso (lookupResult.Kind = LookupResultKind.Good OrElse lookupResult.Kind = LookupResultKind.Inaccessible) Then
                methodGroup = CreateBoundMethodGroup(
                            node,
                            lookupResult,
                            QueryOperatorLookupOptions,
                            source,
                            typeArgumentsOpt,
                            QualificationKind.QualifiedViaValue).MakeCompilerGenerated()
            End If

            diagnostics.Add(node, useSiteDiagnostics)
            lookupResult.Free()

            Return methodGroup
        End Function

    End Class

End Namespace

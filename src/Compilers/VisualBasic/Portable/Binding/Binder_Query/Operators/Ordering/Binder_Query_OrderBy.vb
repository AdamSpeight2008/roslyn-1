' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Order By operator.
        ''' 
        '''     {Preceding query operators} Order By {orderings}
        ''' 
        ''' Ex: From a In AA Order By a ==> AA.OrderBy(Function(a) a)
        ''' 
        ''' Ex: From a In AA Order By a.Key1, a.Key2 Descending ==> AA.OrderBy(Function(a) a.Key1).ThenByDescending(Function(a) a.Key2)
        ''' 
        ''' </summary>
        Private Function BindOrderByClause(
                                            source As BoundQueryClauseBase,
                                            orderBy As OrderByClauseSyntax,
                                            diagnostics As DiagnosticBag
                                          ) As BoundQueryClause

            Dim suppressDiagnostics As DiagnosticBag = Nothing
            Dim callDiagnostics As DiagnosticBag = diagnostics

            Dim lambdaParameterName As String = GetQueryLambdaParameterName(source.RangeVariables)

            Dim sourceOrPreviousOrdering As BoundQueryPart = source
            Dim orderByOrderings As SeparatedSyntaxList(Of OrderingSyntax) = orderBy.Orderings

            If orderByOrderings.Count = 0 Then
                ' Malformed tree.
                Debug.Assert(orderByOrderings.Count > 0, "Malformed syntax tree.")
                Return New BoundQueryClause(orderBy,
                                            BadExpression(orderBy, source, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated(),
                                            source.RangeVariables,
                                            source.CompoundVariableType,
                                            ImmutableArray.Create(Me),
                                            ErrorTypeSymbol.UnknownResultType,
                                            hasErrors:=True)
            End If

            Dim keyBinder As QueryLambdaBinder = Nothing

            For i = 0 To orderByOrderings.Count - 1
                Internal_BindOrderByClause(source, orderBy, diagnostics, suppressDiagnostics, callDiagnostics, lambdaParameterName, sourceOrPreviousOrdering, orderByOrderings, keyBinder, i)
            Next

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Debug.Assert(sourceOrPreviousOrdering IsNot source)
            Debug.Assert(keyBinder IsNot Nothing)

            Return New BoundQueryClause(orderBy,
                                        sourceOrPreviousOrdering,
                                        source.RangeVariables,
                                        source.CompoundVariableType,
                                        ImmutableArray.Create(Of Binder)(keyBinder),
                                        sourceOrPreviousOrdering.Type)
        End Function

        Private Sub Internal_BindOrderByClause(
                                                source As BoundQueryClauseBase,
                                          ByRef orderBy As OrderByClauseSyntax,
                                                diagnostics As DiagnosticBag,
                                          ByRef suppressDiagnostics As DiagnosticBag,
                                          ByRef callDiagnostics As DiagnosticBag,
                                                lambdaParameterName As String,
                                          ByRef sourceOrPreviousOrdering As BoundQueryPart,
                                          ByRef orderByOrderings As SeparatedSyntaxList(Of OrderingSyntax),
                                          ByRef keyBinder As QueryLambdaBinder,
                                                i As Integer
                                              )

            Dim ordering As OrderingSyntax = orderByOrderings(i)

            ' Create LambdaSymbol for the shape of the key lambda.
            Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(lambdaParameterName, 0,
                                                                                       source.CompoundVariableType,
                                                                                       ordering.Expression, source.RangeVariables)

            Dim lambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.OrderingQueryLambda, LambdaUtilities.GetOrderingLambdaBody(ordering)),
                                                          ImmutableArray.Create(param))

            ' Create binder for a key expression.
            keyBinder = New QueryLambdaBinder(lambdaSymbol, source.RangeVariables)

            ' Bind expression as a value, conversion during overload resolution should take care of the rest (making it an RValue, etc.). 
            Dim key As BoundExpression = keyBinder.BindValue(ordering.Expression, diagnostics)

            Dim keyLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                   source.RangeVariables,
                                                   key,
                                                   exprIsOperandOfConditionalBranch:=False)

            keyLambda.SetWasCompilerGenerated()

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If sourceOrPreviousOrdering.Type.IsErrorType() Then
                boundCallOrBadExpression = BadExpression(ordering, ImmutableArray.Create(Of BoundExpression)(sourceOrPreviousOrdering, keyLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                If suppressDiagnostics Is Nothing AndAlso ShouldSuppressDiagnostics(keyLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    callDiagnostics = suppressDiagnostics
                End If

                Dim [operator] = GetMethodNameOfOrderingClause(orderBy, orderByOrderings, i, ordering)

                boundCallOrBadExpression = BindQueryOperatorCall(ordering, sourceOrPreviousOrdering,
                                                               [operator].operatorName,
                                                               ImmutableArray.Create(Of BoundExpression)(keyLambda),
                                                               [operator].operatorNameLocation,
                                                               callDiagnostics)
            End If

            sourceOrPreviousOrdering = New BoundOrdering(ordering, boundCallOrBadExpression, boundCallOrBadExpression.Type)
        End Sub

        Private Shared function GetMethodNameOfOrderingClause(
                                                         ByRef orderBy As OrderByClauseSyntax,
                                                         ByRef orderByOrderings As SeparatedSyntaxList(Of OrderingSyntax),
                                                               i As Integer,
                                                               ordering As OrderingSyntax
                                                             ) As (operatorName As String, operatorNameLocation As TextSpan)
            Dim result As (operatorName As String, operatorNameLocation As TextSpan) = (Nothing, Nothing)
            If i = 0 Then
                ' This is the first ordering.
                If ordering.Kind = SyntaxKind.AscendingOrdering Then
                    result.operatorName = StringConstants.OrderByMethod
                Else
                    Debug.Assert(ordering.Kind = SyntaxKind.DescendingOrdering)
                    result.operatorName = StringConstants.OrderByDescendingMethod
                End If

               result.operatorNameLocation = GetQueryOperatorNameSpan(orderBy.OrderKeyword, orderBy.ByKeyword)
            Else
                If ordering.Kind = SyntaxKind.AscendingOrdering Then
                    result.operatorName = StringConstants.ThenByMethod
                Else
                    Debug.Assert(ordering.Kind = SyntaxKind.DescendingOrdering)
                    result.operatorName = StringConstants.ThenByDescendingMethod
                End If

                result.operatorNameLocation = orderByOrderings.GetSeparator(i - 1).Span
            End If
            Return result
        End function

    End Class

End Namespace

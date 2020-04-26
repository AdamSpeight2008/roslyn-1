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
        ''' Given result of binding preceding query operators, the source, bind the following Let operator.
        ''' 
        '''     {Preceding query operators} Let {expression range variables}
        ''' 
        ''' Ex: From a In AA Let b  ==> AA.Select(Function(a) New With {a, b})
        ''' 
        ''' Ex: From a In AA Let b, c  ==> AA.Select(Function(a) New With {a, b}).Select(Function({a, b}) New With {a, b, c})
        ''' 
        ''' Note, that preceding Select operator can introduce unnamed range variable, which is dropped by the Let
        ''' 
        ''' Ex: From a In AA Select a + 1 Let b ==> AA.Select(Function(a) a + 1).Select(Function(unnamed) b)  
        ''' 
        ''' Also, depending on the amount of expression range variables declared by the Let, and the following query operators,
        ''' translation can produce a nested, as opposed to flat, compound variable.
        ''' 
        ''' Ex: From a In AA Let b, c, d ==> AA.Select(Function(a) New With {a, b}).
        '''                                     Select(Function({a, b}) New With {{a, b}, c}).
        '''                                     Select(Function({{a, b}, c}) New With {a, b, c, d})   
        ''' </summary>
        Private Function BindLetClause(
                                        source As BoundQueryClauseBase,
                                        clauseSyntax As LetClauseSyntax,
                                        operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                        diagnostics As DiagnosticBag,
                               Optional skipFirstVariable As Boolean = False
                                      ) As BoundQueryClause

            Debug.Assert(clauseSyntax Is operatorsEnumerator.Current)

            Dim suppressDiagnostics As DiagnosticBag = Nothing
            Dim callDiagnostics As DiagnosticBag = diagnostics

            Dim variables As SeparatedSyntaxList(Of ExpressionRangeVariableSyntax) = clauseSyntax.Variables
            Debug.Assert(variables.Count > 0, "Malformed syntax tree.")

            If variables.Count = 0 Then
                ' Handle malformed tree gracefully.
                Debug.Assert(Not skipFirstVariable)
                Return New BoundQueryClause(clauseSyntax,
                                            BadExpression(clauseSyntax, source, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated(),
                                            source.RangeVariables.Add(RangeVariableSymbol.CreateForErrorRecovery(Me,
                                                                                                                 clauseSyntax,
                                                                                                                 ErrorTypeSymbol.UnknownResultType)),
                                            ErrorTypeSymbol.UnknownResultType,
                                            ImmutableArray.Create(Me),
                                            ErrorTypeSymbol.UnknownResultType,
                                            hasErrors:=True)
            End If

            Debug.Assert(Not skipFirstVariable OrElse source.Syntax Is variables.First)

            For i = If(skipFirstVariable, 1, 0) To variables.Count - 1

                Dim variable As ExpressionRangeVariableSyntax = variables(i)

                ' Create LambdaSymbol for the shape of the selector lambda.
                Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                           source.CompoundVariableType,
                                                                                           variable, source.RangeVariables)

                Dim lambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.LetVariableQueryLambda,LambdaUtilities.GetLetVariableLambdaBody(variable)),
                                                              ImmutableArray.Create(param))

                ' Create binder for a variable expression.
                Dim selectorBinder As New QueryLambdaBinder(lambdaSymbol, source.RangeVariables)

                Dim declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
                Dim selector As BoundExpression = selectorBinder.BindLetClauseVariableSelector(variable,
                                                                                               operatorsEnumerator,
                                                                                               declaredRangeVariables,
                                                                                               diagnostics)

                Dim selectorLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                            source.RangeVariables,
                                                            selector,
                                                            exprIsOperandOfConditionalBranch:=False)

                lambdaSymbol.SetQueryLambdaReturnType(selector.Type)
                selectorLambda.SetWasCompilerGenerated()

                ' Now bind the call.
                Dim boundCallOrBadExpression As BoundExpression

                If source.Type.IsErrorType() Then
                    boundCallOrBadExpression = BadExpression(variable, ImmutableArray.Create(Of BoundExpression)(source, selectorLambda),
                                                             ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                Else
                    If suppressDiagnostics Is Nothing AndAlso ShouldSuppressDiagnostics(selectorLambda) Then
                        ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                        suppressDiagnostics = DiagnosticBag.GetInstance()
                        callDiagnostics = suppressDiagnostics
                    End If

                    Dim operatorNameLocation As TextSpan

                    If i = 0 Then
                        ' This is the first variable.
                        operatorNameLocation = clauseSyntax.LetKeyword.Span
                    Else
                        operatorNameLocation = variables.GetSeparator(i - 1).Span
                    End If

                    boundCallOrBadExpression = BindQueryOperatorCall(variable, source,
                                                                     StringConstants.SelectMethod,
                                                                     ImmutableArray.Create(Of BoundExpression)(selectorLambda),
                                                                     operatorNameLocation,
                                                                     callDiagnostics)
                End If

                source = New BoundQueryClause(variable,
                                              boundCallOrBadExpression,
                                              source.RangeVariables.Concat(declaredRangeVariables),
                                              selectorLambda.Expression.Type,
                                              ImmutableArray.Create(Of Binder)(selectorBinder),
                                              boundCallOrBadExpression.Type)
            Next

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Return DirectCast(source, BoundQueryClause)
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Apply implicit Select operator at the end of the query to 
        ''' ensure that at least one query operator is called.
        ''' 
        ''' Basically makes query like: 
        '''     From a In AA
        ''' into:
        '''     From a In AA Select a
        ''' </summary>
        Private Function BindFinalImplicitSelectClause(
                                                        source As BoundQueryClauseBase,
                                                        diagnostics As DiagnosticBag
                                                      ) As BoundQueryClause

            Debug.Assert(Not source.Type.IsErrorType())

            Dim fromClauseSyntax = DirectCast(source.Syntax.Parent, FromClauseSyntax)

            ' Create LambdaSymbol for the shape of the selector.
            Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                       source.CompoundVariableType,
                                                                                       fromClauseSyntax, source.RangeVariables)

            ' An implicit selector is an identity (x => x) function that doesn't contain any user code.
            LambdaUtilities.IsNonUserCodeQueryLambda(fromClauseSyntax)
            Dim lambdaSymbol = CreateQueryLambdaSymbol((SynthesizedLambdaKind.FromNonUserCodeQueryLambda,fromClauseSyntax),
                                                          ImmutableArray.Create(param))

            lambdaSymbol.SetQueryLambdaReturnType(source.CompoundVariableType)

            Dim selector As BoundExpression = New BoundParameter(param.Syntax,
                                                                 param,
                                                                 isLValue:=False,
                                                                 type:=param.Type).MakeCompilerGenerated()

            Debug.Assert(Not selector.HasErrors)

            Dim selectorLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                        ImmutableArray(Of RangeVariableSymbol).Empty,
                                                        selector,
                                                        exprIsOperandOfConditionalBranch:=False)
            selectorLambda.SetWasCompilerGenerated()

            Debug.Assert(Not selectorLambda.HasErrors)

            Dim suppressDiagnostics As DiagnosticBag = Nothing

            If param.Type.IsErrorType() Then
                suppressDiagnostics = DiagnosticBag.GetInstance()
                diagnostics = suppressDiagnostics
            End If

            Dim boundCallOrBadExpression As BoundExpression
            boundCallOrBadExpression = BindQueryOperatorCall(source.Syntax.Parent,
                                                             source,
                                                             StringConstants.SelectMethod,
                                                             ImmutableArray.Create(Of BoundExpression)(selectorLambda),
                                                             source.Syntax.Span,
                                                             diagnostics)

            Debug.Assert(boundCallOrBadExpression.WasCompilerGenerated)

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Return New BoundQueryClause(source.Syntax.Parent,
                                        boundCallOrBadExpression,
                                        ImmutableArray(Of RangeVariableSymbol).Empty,
                                        source.CompoundVariableType,
                                        ImmutableArray(Of Binder).Empty,
                                        boundCallOrBadExpression.Type).MakeCompilerGenerated()
        End Function

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Select operator.
        ''' 
        '''     {Preceding query operators} Select {expression range variables}
        ''' 
        ''' From a In AA Select b  ==> AA.Select(Function(a) b)
        ''' 
        ''' From a In AA Select b, c  ==> AA.Select(Function(a) New With {b, c})
        ''' </summary>
        Private Function BindSelectClause(
                                           source As BoundQueryClauseBase,
                                           clauseSyntax As SelectClauseSyntax,
                                           operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                           diagnostics As DiagnosticBag
                                         ) As BoundQueryClause

            Debug.Assert(clauseSyntax Is operatorsEnumerator.Current)

            ' Create LambdaSymbol for the shape of the selector.
            Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                       source.CompoundVariableType,
                                                                                       clauseSyntax, source.RangeVariables)

            Dim lambdaSymbol = CreateQueryLambdaSymbol((SynthesizedLambdaKind.SelectQueryLambda, LambdaUtilities.GetSelectLambdaBody(clauseSyntax)),
                                                        ImmutableArray.Create(param))

            ' Create binder for the selector.
            Dim selectorBinder As New QueryLambdaBinder(lambdaSymbol, source.RangeVariables)

            Dim declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim selector = selectorBinder.BindSelectClauseSelector(clauseSyntax,
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
                boundCallOrBadExpression = BadExpression(clauseSyntax, ImmutableArray.Create(Of BoundExpression)(source, selectorLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If ShouldSuppressDiagnostics(selectorLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If

                boundCallOrBadExpression = BindQueryOperatorCall(clauseSyntax, source,
                                                                 StringConstants.SelectMethod,
                                                                 ImmutableArray.Create(Of BoundExpression)(selectorLambda),
                                                                 clauseSyntax.SelectKeyword.Span,
                                                                 diagnostics)

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            End If

            Return New BoundQueryClause(clauseSyntax,
                                        boundCallOrBadExpression,
                                        declaredRangeVariables,
                                        selectorLambda.Expression.Type,
                                        ImmutableArray.Create(Of Binder)(selectorBinder),
                                        boundCallOrBadExpression.Type)
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        <Conditional("DEBUG")>
        Private Shared Sub DEBUG_Assert_MethodGroupIsNullOrGoodOrInaccessable(methodGroup As BoundMethodGroup)
            Debug.Assert(methodGroup Is Nothing OrElse
                         methodGroup.ResultKind = LookupResultKind.Good OrElse
                         methodGroup.ResultKind = LookupResultKind.Inaccessible)
        End Sub

        ''' <summary>
        ''' Infer type of the group for a Group By operator from the set of available GroupBy methods.
        ''' 
        ''' In short, given already bound itemsLambda and keysLambda, this method performs overload
        ''' resolution over the set of available GroupBy operator methods using fake Into lambda:
        '''     Function(key, group As typeToBeInferred) New With {group}
        ''' 
        ''' If resolution succeeds, the type inferred for the best candidate is our result.  
        ''' </summary>
        Private Function InferGroupType(
                                         source As BoundQueryClauseBase,
                                         groupBy As GroupByClauseSyntax,
                                         itemsLambda As BoundQueryLambda,
                                         keysLambda As BoundQueryLambda,
                                         keysRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                             <Out> ByRef methodGroup As BoundMethodGroup,
                                         diagnostics As DiagnosticBag
                                       ) As TypeSymbol

            Debug.Assert(methodGroup Is Nothing)

            Dim groupType As TypeSymbol = ErrorTypeSymbol.UnknownResultType

            If Not source.Type.IsErrorType() Then

                methodGroup = LookupQueryOperator(groupBy, source, StringConstants.GroupByMethod, Nothing, diagnostics)
                DEBUG_Assert_MethodGroupIsNullOrGoodOrInaccessable(methodGroup)

                If methodGroup Is Nothing Then
                    ReportDiagnostic(diagnostics, Location.Create(groupBy.SyntaxTree, GetGroupByOperatorNameSpan(groupBy)), ERRID.ERR_QueryOperatorNotFound, StringConstants.GroupByMethod)

                ElseIf Not (ShouldSuppressDiagnostics(keysLambda) OrElse
                         (itemsLambda IsNot Nothing AndAlso ShouldSuppressDiagnostics(itemsLambda))) Then

                    Dim inferenceLambda As New GroupTypeInferenceLambda(groupBy, Me,
                                                                        (New ParameterSymbol() {
                                                                            CreateQueryLambdaParameterSymbol(StringConstants.It1, 0, keysLambda.Expression.Type, groupBy, keysRangeVariables),
                                                                            CreateQueryLambdaParameterSymbol(StringConstants.It2, 1, Nothing, groupBy)}).AsImmutableOrNull(),
                                                                        Compilation)

                    Dim groupByArguments() As BoundExpression

                    If itemsLambda Is Nothing Then
                        groupByArguments = {keysLambda, inferenceLambda}
                    Else
                        groupByArguments = {keysLambda, itemsLambda, inferenceLambda}
                    End If

                    Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
                    Dim results As OverloadResolution.OverloadResolutionResult = OverloadResolution.QueryOperatorInvocationOverloadResolution(methodGroup,
                                                                                                                                              groupByArguments.AsImmutableOrNull(), Me,
                                                                                                                                              useSiteDiagnostics)

                    diagnostics.Add(groupBy, useSiteDiagnostics)

                    If results.BestResult.HasValue Then
                        Dim method = DirectCast(results.BestResult.Value.Candidate.UnderlyingSymbol, MethodSymbol)
                        Dim resultSelector As TypeSymbol = method.Parameters(groupByArguments.Length - 1).Type

                        groupType = resultSelector.DelegateOrExpressionDelegate(Me).DelegateInvokeMethod.Parameters(1).Type

                    ElseIf Not source.HasErrors Then
                        ReportDiagnostic(diagnostics, Location.Create(groupBy.SyntaxTree, GetGroupByOperatorNameSpan(groupBy)), ERRID.ERR_QueryOperatorNotFound, StringConstants.GroupByMethod)
                    End If
                End If
            End If

            Debug.Assert(groupType IsNot Nothing)
            Return groupType
        End Function

        ''' <summary>
        ''' Infer type of the group for a Group Join operator from the set of available GroupJoin methods.
        ''' 
        ''' In short, given already bound inner source and the join key lambdas, this method performs overload
        ''' resolution over the set of available GroupJoin operator methods using fake Into lambda:
        '''     Function(outerVar, group As typeToBeInferred) New With {group}
        ''' 
        ''' If resolution succeeds, the type inferred for the best candidate is our result.  
        ''' </summary>
        Private Function InferGroupType(
                                         outer As BoundQueryClauseBase,
                                         inner As BoundQueryClauseBase,
                                         groupJoin As GroupJoinClauseSyntax,
                                         outerKeyLambda As BoundQueryLambda,
                                         innerKeyLambda As BoundQueryLambda,
                             <Out> ByRef methodGroup As BoundMethodGroup,
                                         diagnostics As DiagnosticBag
                                       ) As TypeSymbol

            Debug.Assert(methodGroup Is Nothing)

            Dim groupType As TypeSymbol = ErrorTypeSymbol.UnknownResultType

            If Not outer.Type.IsErrorType() Then

                ' If outer's type is "good", we should always do a lookup, even if we won't attempt the inference
                ' because BindGroupJoinClause still expects to have a group, unless the lookup fails.
                methodGroup = LookupQueryOperator(groupJoin, outer, StringConstants.GroupJoinMethod, Nothing, diagnostics)
                DEBUG_Assert_MethodGroupIsNullOrGoodOrInaccessable(methodGroup)

                If methodGroup Is Nothing Then
                    ReportDiagnostic(diagnostics,
                                     Location.Create(groupJoin.SyntaxTree, GetQueryOperatorNameSpan(groupJoin.GroupKeyword, groupJoin.JoinKeyword)),
                                     ERRID.ERR_QueryOperatorNotFound, StringConstants.GroupJoinMethod)

                ElseIf Not ShouldSuppressDiagnostics(innerKeyLambda) AndAlso Not ShouldSuppressDiagnostics(outerKeyLambda) AndAlso
                       Not inner.HasErrors AndAlso Not inner.Type.IsErrorType() Then

                    Dim inferenceLambda As New GroupTypeInferenceLambda(groupJoin, Me,
                                                                        (New ParameterSymbol() {
                                    CreateQueryLambdaParameterSymbol(StringConstants.It1,
                                                                   0,
                                                                   outer.CompoundVariableType,
                                                                   groupJoin, outer.RangeVariables),
                                    CreateQueryLambdaParameterSymbol(StringConstants.It2,
                                                                   1,
                                                                   Nothing,
                                                                   groupJoin)}).AsImmutableOrNull(),
                                                                        Compilation)

                    Dim groupJoinArguments() As BoundExpression = {inner, outerKeyLambda, innerKeyLambda, inferenceLambda}

                    Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
                    Dim results As OverloadResolution.OverloadResolutionResult = OverloadResolution.QueryOperatorInvocationOverloadResolution(methodGroup,
                                                                                                                                              groupJoinArguments.AsImmutableOrNull(), Me,
                                                                                                                                              useSiteDiagnostics)

                    diagnostics.Add(groupJoin, useSiteDiagnostics)

                    If results.BestResult.HasValue Then
                        Dim method = DirectCast(results.BestResult.Value.Candidate.UnderlyingSymbol, MethodSymbol)
                        Dim resultSelector As TypeSymbol = method.Parameters(groupJoinArguments.Length - 1).Type

                        Dim resultSelectorDelegate = resultSelector.DelegateOrExpressionDelegate(Me)
                        groupType = resultSelectorDelegate.DelegateInvokeMethod.Parameters(1).Type

                    ElseIf Not outer.HasErrors Then
                        ReportDiagnostic(diagnostics,
                                         Location.Create(groupJoin.SyntaxTree, GetQueryOperatorNameSpan(groupJoin.GroupKeyword, groupJoin.JoinKeyword)),
                                         ERRID.ERR_QueryOperatorNotFound, StringConstants.GroupJoinMethod)
                    End If
                End If
            End If

            Debug.Assert(groupType IsNot Nothing)
            Return groupType
        End Function

        Friend Overridable Function BindGroupAggregationExpression(
                                                                    group As GroupAggregationSyntax,
                                                                    diagnostics As DiagnosticBag
                                                                  ) As BoundExpression
            ' Only special query binders that have enough context can bind GroupAggregationSyntax.
            ' TODO: Do we need to report any diagnostic?
            Debug.Assert(False, "Binding out of context is unsupported!")
            Return BadExpression(group, ErrorTypeSymbol.UnknownResultType)
        End Function

    End Class

End Namespace

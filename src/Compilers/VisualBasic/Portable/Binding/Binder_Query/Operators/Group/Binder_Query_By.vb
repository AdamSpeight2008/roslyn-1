' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Group By operator.
        ''' 
        '''     [{Preceding query operators}] Group [{items expression range variables}] 
        '''                                   By {keys expression range variables}
        '''                                   Into {aggregation range variables}
        ''' 
        ''' Ex: From a In AA Group By Key(a)          AA.GroupBy(Function(a) Key(a), 
        '''                  Into Count()     ==>                Function(key, group_a) New With {key, group_a.Count()})
        '''                  
        ''' Ex: From a In AA Group Item(a)            AA.GroupBy(Function(a) Key(a), 
        '''                  By Key(a)        ==>                Function(a) Item(a), 
        '''                  Into Count()                        Function(key, group_a) New With {key, group_a.Count()})
        ''' 
        ''' Note, that type of the group must be inferred from the set of available GroupBy operators in order to be able to 
        ''' interpret the aggregation range variables. 
        ''' </summary>
        Private Function BindGroupByClause(
                                            source As BoundQueryClauseBase,
                                            groupBy As GroupByClauseSyntax,
                                            diagnostics As DiagnosticBag
                                          ) As BoundQueryClause

            ' Handle group items.
            Dim itemsLambdaBinder As QueryLambdaBinder = Nothing
            Dim itemsRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim itemsLambda As BoundQueryLambda = BindGroupByItems(source, groupBy, itemsLambdaBinder, itemsRangeVariables, diagnostics)

            ' Handle grouping keys.
            Dim keysLambdaBinder As QueryLambdaBinder = Nothing
            Dim keysRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim keysLambda As BoundQueryLambda = BindGroupByKeys(source, groupBy, keysLambdaBinder, keysRangeVariables, diagnostics)
            Debug.Assert(keysLambda IsNot Nothing)

            ' Infer type of the resulting group.
            Dim methodGroup As BoundMethodGroup = Nothing
            Dim groupType As TypeSymbol = InferGroupType(source, groupBy, itemsLambda, keysLambda, keysRangeVariables, methodGroup, diagnostics)

            ' Bind the INTO selector.
            Dim groupRangeVariables As ImmutableArray(Of RangeVariableSymbol)
            Dim groupCompoundVariableType As TypeSymbol

            If itemsLambda Is Nothing Then
                groupRangeVariables = source.RangeVariables
                groupCompoundVariableType = source.CompoundVariableType
            Else
                groupRangeVariables = itemsRangeVariables
                groupCompoundVariableType = itemsLambda.Expression.Type
            End If

            Dim intoBinder As IntoClauseBinder = Nothing
            Dim intoRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim intoLambda As BoundQueryLambda = BindIntoSelectorLambda(groupBy, keysRangeVariables, keysLambda.Expression.Type, False, Nothing,
                                                                        groupType, groupRangeVariables, groupCompoundVariableType,
                                                                        groupBy.AggregationVariables, True,
                                                                        diagnostics, intoBinder, intoRangeVariables)

            ' Now bind the call.
            Dim groupByArguments() As BoundExpression
            Dim lambdaBinders As ImmutableArray(Of Binder)

            Debug.Assert((itemsLambda Is Nothing) = (itemsLambdaBinder Is Nothing))

            If itemsLambda Is Nothing Then
                groupByArguments = {keysLambda, intoLambda}
                lambdaBinders = ImmutableArray.Create(Of Binder)(keysLambdaBinder, intoBinder)
            Else
                groupByArguments = {keysLambda, itemsLambda, intoLambda}
                lambdaBinders = ImmutableArray.Create(Of Binder)(keysLambdaBinder, itemsLambdaBinder, intoBinder)
            End If

            Dim boundCallOrBadExpression As BoundExpression

            If source.Type.IsErrorType() OrElse methodGroup Is Nothing Then
                boundCallOrBadExpression = BadExpression(groupBy,
                                                         ImmutableArray.Create(Of BoundExpression)(source).AddRange(groupByArguments),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim callDiagnostics As DiagnosticBag = diagnostics

                If ShouldSuppressDiagnostics(keysLambda) OrElse ShouldSuppressDiagnostics(intoLambda) OrElse
                   (itemsLambda IsNot Nothing AndAlso ShouldSuppressDiagnostics(itemsLambda)) Then

                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    callDiagnostics = DiagnosticBag.GetInstance()
                End If

                ' Reusing method group that we got while inferring group type, this way we can avoid doing name lookup again. 
                boundCallOrBadExpression = BindQueryOperatorCall(groupBy, source,
                                                               StringConstants.GroupByMethod,
                                                               methodGroup,
                                                               groupByArguments.AsImmutableOrNull(),
                                                               GetGroupByOperatorNameSpan(groupBy),
                                                               callDiagnostics)

                If callDiagnostics IsNot diagnostics Then
                    callDiagnostics.Free()
                End If
            End If

            Return New BoundQueryClause(groupBy,
                                        boundCallOrBadExpression,
                                        keysRangeVariables.Concat(intoRangeVariables),
                                        intoLambda.Expression.Type,
                                        lambdaBinders,
                                        boundCallOrBadExpression.Type)
        End Function

        Private Shared Function GetGroupByOperatorNameSpan(groupBy As GroupByClauseSyntax) As TextSpan
            If groupBy.Items.Count = 0 Then Return GetQueryOperatorNameSpan(groupBy.GroupKeyword, groupBy.ByKeyword)
            Return groupBy.GroupKeyword.Span
        End Function

        ''' <summary>
        ''' Returns Nothing if items were omitted.
        ''' </summary>
        Private Function BindGroupByItems(
                                           source As BoundQueryClauseBase,
                                           groupBy As GroupByClauseSyntax,
                               <Out> ByRef itemsLambdaBinder As QueryLambdaBinder,
                               <Out> ByRef itemsRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                           diagnostics As DiagnosticBag
                                         ) As BoundQueryLambda

            Debug.Assert(itemsLambdaBinder Is Nothing)
            Debug.Assert(itemsRangeVariables.IsDefault)

            ' Handle group items.
            Dim itemsLambda As BoundQueryLambda = Nothing
            Dim items As SeparatedSyntaxList(Of ExpressionRangeVariableSyntax) = groupBy.Items

            If items.Count > 0 Then
                Dim itemsParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                                source.CompoundVariableType,
                                                                                                groupBy, source.RangeVariables)

                Dim itemsLambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.GroupByItemsQueryLambda,LambdaUtilities.GetGroupByItemsLambdaBody(groupBy)),
                                                                   ImmutableArray.Create(itemsParam))

                ' Create binder for the selector.
                itemsLambdaBinder = New QueryLambdaBinder(itemsLambdaSymbol, source.RangeVariables)

                Dim itemsSelector = itemsLambdaBinder.BindExpressionRangeVariables(items, False, groupBy,
                                                                                   itemsRangeVariables, diagnostics)

                itemsLambda = CreateBoundQueryLambda(itemsLambdaSymbol,
                                                     source.RangeVariables,
                                                     itemsSelector,
                                                     exprIsOperandOfConditionalBranch:=False)

                itemsLambdaSymbol.SetQueryLambdaReturnType(itemsSelector.Type)
                itemsLambda.SetWasCompilerGenerated()
            Else
                itemsLambdaBinder = Nothing
                itemsRangeVariables = ImmutableArray(Of RangeVariableSymbol).Empty
            End If

            Debug.Assert((itemsLambda Is Nothing) = (itemsLambdaBinder Is Nothing))
            Debug.Assert(Not itemsRangeVariables.IsDefault)
            Debug.Assert(itemsLambda IsNot Nothing OrElse itemsRangeVariables.Length = 0)
            Return itemsLambda
        End Function

        Private Function BindGroupByKeys(
                                          source As BoundQueryClauseBase,
                                          groupBy As GroupByClauseSyntax,
                              <Out> ByRef keysLambdaBinder As QueryLambdaBinder,
                              <Out> ByRef keysRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                          diagnostics As DiagnosticBag
                                        ) As BoundQueryLambda

            Debug.Assert(keysLambdaBinder Is Nothing)
            Debug.Assert(keysRangeVariables.IsDefault)

            Dim keys As SeparatedSyntaxList(Of ExpressionRangeVariableSyntax) = groupBy.Keys

            Dim keysParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                           source.CompoundVariableType,
                                                                                           groupBy, source.RangeVariables)

            Dim keysLambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.GroupByKeysQueryLambda,LambdaUtilities.GetGroupByKeysLambdaBody(groupBy)),
                                                              ImmutableArray.Create(keysParam))

            ' Create binder for the selector.
            keysLambdaBinder = New QueryLambdaBinder(keysLambdaSymbol, source.RangeVariables)

            Dim keysSelector = keysLambdaBinder.BindExpressionRangeVariables(keys, True, groupBy,
                                                                             keysRangeVariables, diagnostics)

            Dim keysLambda = CreateBoundQueryLambda(keysLambdaSymbol,
                                                    source.RangeVariables,
                                                    keysSelector,
                                                    exprIsOperandOfConditionalBranch:=False)

            keysLambdaSymbol.SetQueryLambdaReturnType(keysSelector.Type)
            keysLambda.SetWasCompilerGenerated()

            Return keysLambda
        End Function

    End Class

End Namespace

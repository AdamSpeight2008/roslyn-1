' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Bind query expression that starts with Aggregate keyword, as opposed to the one that starts with From.
        ''' 
        '''     Aggregate {collection range variables} [{other operators}] Into {aggregation range variables}
        ''' 
        ''' If Into clause has one item, a single value is produced. If it has multiple items, values are
        ''' combined into an instance of an Anonymous Type.
        ''' </summary>
        Private Function BindAggregateQueryExpression(
                                                       query As QueryExpressionSyntax,
                                                       operators As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                       diagnostics As DiagnosticBag
                                                     ) As BoundQueryExpression

            Dim aggregate = DirectCast(operators.Current, AggregateClauseSyntax)

            Dim malformedSyntax As Boolean = operators.MoveNext()

            Debug.Assert(Not malformedSyntax, "Malformed syntax tree. Parser shouldn't produce a tree like that.")

            operators = aggregate.AdditionalQueryOperators.GetEnumerator()

            ' Note, this call can advance [operators] enumerator if it absorbs the following Let or Select.
            Dim source As BoundQueryClauseBase = BindCollectionRangeVariables(aggregate, Nothing, aggregate.Variables, operators, diagnostics)
            source = BindSubsequentQueryOperators(source, operators, diagnostics)

            Dim aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax) = aggregate.AggregationVariables
            Dim aggregationVariablesCount As Integer = aggregationVariables.Count

            Select Case aggregationVariablesCount
                Case 0
                    source = AggregateZeroVariables(aggregate, source, aggregationVariables)

                Case 1
                    ' Simple case, one item in the [Into] clause, source is our group.
                    source = AggregateOneVariables(diagnostics, aggregate, source, aggregationVariables)

                Case Else
                    ' Complex case, need to build an instance of an Anonymous Type.
                    source = AggregateComplexCase(diagnostics, aggregate, source, aggregationVariables, aggregationVariablesCount)
            End Select

            Debug.Assert(Not source.Binders.IsDefault AndAlso source.Binders.Length = 1 AndAlso source.Binders(0) IsNot Nothing)

            Return New BoundQueryExpression(query, source, If(malformedSyntax, ErrorTypeSymbol.UnknownResultType, source.Type), hasErrors:=malformedSyntax)
        End Function

        Private Function AggregateComplexCase(
                                               diagnostics As DiagnosticBag,
                                               aggregate As AggregateClauseSyntax,
                                               source As BoundQueryClauseBase,
                                               aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax),
                                               aggregationVariablesCount As Integer
                                             ) As BoundQueryClauseBase
            ' Complex case, need to build an instance of an Anonymous Type. 
            Dim declaredNames As PooledHashSet(Of String) = CreateSetOfDeclaredNames()

            Dim selectors = New BoundExpression(aggregationVariablesCount - 1) {}
            Dim fields = New AnonymousTypeField(selectors.Length - 1) {}

            Dim groupReference = New BoundRValuePlaceholder(aggregate, source.Type).MakeCompilerGenerated()
            Dim intoBinder As New IntoClauseDisallowGroupReferenceBinder(Me, groupReference, source.RangeVariables, source.CompoundVariableType, source.RangeVariables)

            For i As Integer = 0 To aggregationVariablesCount - 1
                Dim rangeVarSymbol = intoBinder.BindAggregationRangeVariable(aggregationVariables(i),
                                                                             declaredNames,
                                                                             selectors(i),
                                                                             diagnostics)

                Debug.Assert(rangeVarSymbol IsNot Nothing)
                fields(i) = New AnonymousTypeField(rangeVarSymbol.Name, rangeVarSymbol.Type, rangeVarSymbol.Syntax.GetLocation(), isKeyOrByRef:=True)
            Next

            Dim result As BoundExpression = BindAnonymousObjectCreationExpression(aggregate,
                                                     New AnonymousTypeDescriptor(fields.AsImmutableOrNull(),
                                                                                 aggregate.IntoKeyword.GetLocation(),
                                                                                 True),
                                                     selectors.AsImmutableOrNull(),
                                                     diagnostics).MakeCompilerGenerated()

            source = New BoundAggregateClause(aggregate, source, groupReference,
                                              result,
                                              ImmutableArray(Of RangeVariableSymbol).Empty,
                                              result.Type,
                                              ImmutableArray.Create(Of Binder)(intoBinder),
                                              result.Type)
            Return source
        End Function

        Private Function AggregateOneVariables(
                                                diagnostics As DiagnosticBag,
                                                aggregate As AggregateClauseSyntax,
                                                source As BoundQueryClauseBase,
                                                aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax)
                                              ) As BoundQueryClauseBase

            Dim intoBinder As New IntoClauseDisallowGroupReferenceBinder(Me, source, source.RangeVariables, source.CompoundVariableType, source.RangeVariables)

            Dim aggregationSelector As BoundExpression = Nothing

            intoBinder.BindAggregationRangeVariable(aggregationVariables(0),
                                                    Nothing,
                                                    aggregationSelector,
                                                    diagnostics)

            source = New BoundAggregateClause(aggregate, Nothing, Nothing,
                                              aggregationSelector,
                                              ImmutableArray(Of RangeVariableSymbol).Empty,
                                              aggregationSelector.Type,
                                              ImmutableArray.Create(Of Binder)(intoBinder),
                                              aggregationSelector.Type)
            Return source
        End Function

        Private Function AggregateZeroVariables(
                                                 aggregate As AggregateClauseSyntax,
                                                 source As BoundQueryClauseBase,
                                                 aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax)
                                               ) As BoundQueryClauseBase

            Debug.Assert(aggregationVariables.Count > 0, "Malformed syntax tree.")

            Dim intoBinder As New IntoClauseDisallowGroupReferenceBinder(Me, source, source.RangeVariables, source.CompoundVariableType, source.RangeVariables)

            source = New BoundAggregateClause(aggregate, Nothing, Nothing,
                                              BadExpression(aggregate, source, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated(),
                                              ImmutableArray(Of RangeVariableSymbol).Empty,
                                              ErrorTypeSymbol.UnknownResultType,
                                              ImmutableArray.Create(Of Binder)(intoBinder),
                                              ErrorTypeSymbol.UnknownResultType)
            Return source
        End Function

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Aggregate operator.
        ''' 
        '''     {Preceding query operators} Aggregate {collection range variables} [{other operators}] Into {aggregation range variables}
        ''' 
        ''' Depending on how many items we have in the INTO clause,
        ''' we will interpret Aggregate operator as follows:
        '''
        ''' FROM a in AA              FROM a in AA
        ''' AGGREGATE b in a.BB  =>   LET count = (FROM b IN a.BB).Count()
        ''' INTO Count()
        '''
        ''' FROM a in AA              FROM a in AA
        ''' AGGREGATE b in a.BB  =>   LET Group = (FROM b IN a.BB)
        ''' INTO Count(),             Select a, Count=Group.Count(), Sum=Group.Sum(b=>b)
        '''      Sum(b)
        '''
        ''' </summary>
        Private Function BindAggregateClause(
                                              source As BoundQueryClauseBase,
                                              aggregate As AggregateClauseSyntax,
                                              operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                              diagnostics As DiagnosticBag
                                            ) As BoundAggregateClause

            Debug.Assert(operatorsEnumerator.Current Is aggregate)

            ' Let's interpret our group.
            ' Create LambdaSymbol for the shape of the Let-selector lambda.
            Dim letSelectorParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(
                                                                    GetQueryLambdaParameterName(source.RangeVariables),
                                                                    0,
                                                                    source.CompoundVariableType,
                                                                    aggregate,
                                                                    source.RangeVariables)

            ' The lambda that will contain the nested query.
            Dim letSelectorLambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.AggregateQueryLambda,LambdaUtilities.GetAggregateLambdaBody(aggregate)),
                                                                     ImmutableArray.Create(letSelectorParam))

            ' Create binder for the [Let] selector.
            Dim letSelectorBinder As New QueryLambdaBinder(letSelectorLambdaSymbol, source.RangeVariables)

            Dim declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim group As BoundQueryClauseBase = Nothing
            Dim intoBinder As IntoClauseDisallowGroupReferenceBinder = Nothing
            Dim letSelector As BoundExpression = letSelectorBinder.BindAggregateClauseFirstSelector(aggregate, operatorsEnumerator,
                                                                                                    source.RangeVariables,
                                                                                                    ImmutableArray(Of RangeVariableSymbol).Empty,
                                                                                                    declaredRangeVariables,
                                                                                                    group,
                                                                                                    intoBinder,
                                                                                                    diagnostics)

            Dim letSelectorLambda As BoundQueryLambda

            letSelectorLambda = CreateBoundQueryLambda(letSelectorLambdaSymbol,
                                                       source.RangeVariables,
                                                       letSelector,
                                                       exprIsOperandOfConditionalBranch:=False)

            letSelectorLambdaSymbol.SetQueryLambdaReturnType(letSelector.Type)
            letSelectorLambda.SetWasCompilerGenerated()


            ' Now bind the [Let] operator call.
            Dim suppressDiagnostics As DiagnosticBag = Nothing
            Dim underlyingExpression As BoundExpression

            If source.Type.IsErrorType() Then
                underlyingExpression = BadExpression(aggregate, ImmutableArray.Create(Of BoundExpression)(source, letSelectorLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim callDiagnostics As DiagnosticBag = diagnostics

                If ShouldSuppressDiagnostics(letSelectorLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    Debug.Assert(suppressDiagnostics Is Nothing)
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    callDiagnostics = suppressDiagnostics
                End If

                underlyingExpression = BindQueryOperatorCall(aggregate, source,
                                                             StringConstants.SelectMethod,
                                                             ImmutableArray.Create(Of BoundExpression)(letSelectorLambda),
                                                             aggregate.AggregateKeyword.Span,
                                                             callDiagnostics)
            End If

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Return CompleteAggregateClauseBinding(aggregate,
                                                  operatorsEnumerator,
                                                  source.RangeVariables,
                                                  ImmutableArray(Of RangeVariableSymbol).Empty,
                                                  underlyingExpression,
                                                  letSelectorBinder,
                                                  declaredRangeVariables,
                                                  letSelectorLambda.Expression.Type,
                                                  group,
                                                  intoBinder,
                                                  diagnostics)
        End Function

        Private Function CompleteAggregateClauseBinding(
                                                         aggregate As AggregateClauseSyntax,
                                                         operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                         sourceRangeVariablesPart1 As ImmutableArray(Of RangeVariableSymbol),
                                                         sourceRangeVariablesPart2 As ImmutableArray(Of RangeVariableSymbol),
                                                         firstSelectExpression As BoundExpression,
                                                         firstSelectSelectorBinder As QueryLambdaBinder,
                                                         firstSelectDeclaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                         firstSelectCompoundVariableType As TypeSymbol,
                                                         group As BoundQueryClauseBase,
                                                         intoBinder As IntoClauseDisallowGroupReferenceBinder,
                                                         diagnostics As DiagnosticBag
                                                       ) As BoundAggregateClause

            Debug.Assert((sourceRangeVariablesPart1.Length = 0) = (sourceRangeVariablesPart2 = firstSelectSelectorBinder.RangeVariables))
            Debug.Assert((sourceRangeVariablesPart2.Length = 0) = (sourceRangeVariablesPart1 = firstSelectSelectorBinder.RangeVariables))
            Debug.Assert(firstSelectSelectorBinder.RangeVariables.Length = sourceRangeVariablesPart1.Length + sourceRangeVariablesPart2.Length)

            Dim result As BoundAggregateClause
            Dim aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax) = aggregate.AggregationVariables

            If aggregationVariables.Count <= 1 Then
                ' Simple case
                Debug.Assert(intoBinder IsNot Nothing)
                Debug.Assert(firstSelectDeclaredRangeVariables.Length <= 1)
                result = New BoundAggregateClause(aggregate, Nothing, Nothing,
                                                  firstSelectExpression,
                                                  firstSelectSelectorBinder.RangeVariables.Concat(firstSelectDeclaredRangeVariables),
                                                  firstSelectCompoundVariableType,
                                                  ImmutableArray.Create(Of Binder)(firstSelectSelectorBinder, intoBinder),
                                                  firstSelectExpression.Type)
            Else

                ' Complex case, apply the [Select].
                Debug.Assert(intoBinder Is Nothing)
                Debug.Assert(firstSelectDeclaredRangeVariables.Length = 1)

                Dim suppressCallDiagnostics As Boolean = (firstSelectExpression.Kind = BoundKind.BadExpression)

                If Not suppressCallDiagnostics AndAlso firstSelectExpression.HasErrors AndAlso firstSelectExpression.Kind = BoundKind.QueryClause Then
                    Dim query = DirectCast(firstSelectExpression, BoundQueryClause)
                    suppressCallDiagnostics = query.UnderlyingExpression.Kind = BoundKind.BadExpression
                End If

                Dim letOperator = New BoundQueryClause(aggregate,
                                                       firstSelectExpression,
                                                       firstSelectSelectorBinder.RangeVariables.Concat(firstSelectDeclaredRangeVariables),
                                                       firstSelectCompoundVariableType,
                                                       ImmutableArray(Of Binder).Empty,
                                                       firstSelectExpression.Type).MakeCompilerGenerated()

                Dim selectSelectorParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(letOperator.RangeVariables), 0,
                                                                                                         letOperator.CompoundVariableType,
                                                                                                         aggregate, letOperator.RangeVariables)

                ' This lambda only contains non-user code. We associate it with the aggregate clause itself.
                Debug.Assert(LambdaUtilities.IsNonUserCodeQueryLambda(aggregate))
                Dim selectSelectorLambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.AggregateNonUserCodeQueryLambda,aggregate),
                                                                            ImmutableArray.Create(selectSelectorParam))

                ' Create new binder for the [Select] selector.
                Dim groupRangeVar As RangeVariableSymbol = firstSelectDeclaredRangeVariables(0)
                Dim selectSelectorBinder As New QueryLambdaBinder(selectSelectorLambdaSymbol, ImmutableArray(Of RangeVariableSymbol).Empty)
                Dim groupReference = New BoundRangeVariable(groupRangeVar.Syntax, groupRangeVar, groupRangeVar.Type).MakeCompilerGenerated()
                intoBinder = New IntoClauseDisallowGroupReferenceBinder(selectSelectorBinder,
                                                                        groupReference,
                                                                        group.RangeVariables,
                                                                        group.CompoundVariableType,
                                                                        firstSelectSelectorBinder.RangeVariables.Concat(group.RangeVariables)
                                                                       )


                ' Compound range variable after the first [Let] has shape { [<compound key part1>, ][<compound key part2>, ]<group> }.
                Dim compoundKeyReferencePart1 As BoundExpression, keysRangeVariablesPart1 As ImmutableArray(Of RangeVariableSymbol)
                Dim compoundKeyReferencePart2 As BoundExpression, keysRangeVariablesPart2 As ImmutableArray(Of RangeVariableSymbol)

                If sourceRangeVariablesPart1.Length > 0 Then
                    ' So we need to get a reference to the first property of selector's parameter.
                    Dim anonymousType = DirectCast(selectSelectorParam.Type, AnonymousTypeManager.AnonymousTypePublicSymbol)
                    Dim keyProperty = anonymousType.Properties(0)

                    Debug.Assert(keyProperty.Type Is firstSelectSelectorBinder.LambdaSymbol.Parameters(0).Type)

                    compoundKeyReferencePart1 = New BoundPropertyAccess(aggregate,
                                                                        keyProperty,
                                                                        Nothing,
                                                                        PropertyAccessKind.Get,
                                                                        False,
                                                                        New BoundParameter(selectSelectorParam.Syntax,
                                                                                           selectSelectorParam, False,
                                                                                           selectSelectorParam.Type).MakeCompilerGenerated(),
                                                                        ImmutableArray(Of BoundExpression).Empty).MakeCompilerGenerated()

                    keysRangeVariablesPart1 = sourceRangeVariablesPart1

                    If sourceRangeVariablesPart2.Length > 0 Then
                        keyProperty = anonymousType.Properties(1)

                        Debug.Assert(keyProperty.Type Is firstSelectSelectorBinder.LambdaSymbol.Parameters(1).Type)

                        ' We need to get a reference to the second property of selector's parameter.
                        compoundKeyReferencePart2 = New BoundPropertyAccess(aggregate,
                                                                            keyProperty,
                                                                            Nothing,
                                                                            PropertyAccessKind.Get,
                                                                            False,
                                                                            New BoundParameter(selectSelectorParam.Syntax,
                                                                                               selectSelectorParam, False,
                                                                                               selectSelectorParam.Type).MakeCompilerGenerated(),
                                                                            ImmutableArray(Of BoundExpression).Empty).MakeCompilerGenerated()

                        keysRangeVariablesPart2 = sourceRangeVariablesPart2
                    Else
                        compoundKeyReferencePart2 = Nothing
                        keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                    End If
                ElseIf sourceRangeVariablesPart2.Length > 0 Then
                    ' So we need to get a reference to the first property of selector's parameter.
                    Dim anonymousType = DirectCast(selectSelectorParam.Type, AnonymousTypeManager.AnonymousTypePublicSymbol)
                    Dim keyProperty = anonymousType.Properties(0)

                    Debug.Assert(keyProperty.Type Is firstSelectSelectorBinder.LambdaSymbol.Parameters(1).Type)

                    compoundKeyReferencePart1 = New BoundPropertyAccess(aggregate,
                                                                        keyProperty,
                                                                        Nothing,
                                                                        PropertyAccessKind.Get,
                                                                        False,
                                                                        New BoundParameter(selectSelectorParam.Syntax,
                                                                                           selectSelectorParam, False,
                                                                                           selectSelectorParam.Type).MakeCompilerGenerated(),
                                                                        ImmutableArray(Of BoundExpression).Empty).MakeCompilerGenerated()

                    keysRangeVariablesPart1 = sourceRangeVariablesPart2
                    compoundKeyReferencePart2 = Nothing
                    keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                Else
                    compoundKeyReferencePart1 = Nothing : keysRangeVariablesPart1 = ImmutableArray(Of RangeVariableSymbol).Empty
                    compoundKeyReferencePart2 = Nothing : keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                End If

                Dim declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
                Dim selectSelector As BoundExpression = intoBinder.BindIntoSelector(aggregate,
                                                                                    firstSelectSelectorBinder.RangeVariables,
                                                                                    compoundKeyReferencePart1,
                                                                                    keysRangeVariablesPart1,
                                                                                    compoundKeyReferencePart2,
                                                                                    keysRangeVariablesPart2,
                                                                                    Nothing,
                                                                                    aggregationVariables,
                                                                                    MustProduceFlatCompoundVariable(operatorsEnumerator),
                                                                                    declaredRangeVariables,
                                                                                    diagnostics)

                Dim selectSelectorLambda = CreateBoundQueryLambda(selectSelectorLambdaSymbol,
                                                                  letOperator.RangeVariables,
                                                                  selectSelector,
                                                                  exprIsOperandOfConditionalBranch:=False)

                selectSelectorLambdaSymbol.SetQueryLambdaReturnType(selectSelector.Type)
                selectSelectorLambda.SetWasCompilerGenerated()

                Dim underlyingExpression As BoundExpression
                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If letOperator.Type.IsErrorType() Then
                    underlyingExpression = BadExpression(aggregate, ImmutableArray.Create(Of BoundExpression)(letOperator, selectSelectorLambda),
                                                             ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                Else
                    Dim callDiagnostics As DiagnosticBag = diagnostics

                    If suppressCallDiagnostics OrElse ShouldSuppressDiagnostics(selectSelectorLambda) Then
                        ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                        suppressDiagnostics = DiagnosticBag.GetInstance()
                        callDiagnostics = suppressDiagnostics
                    End If

                    underlyingExpression = BindQueryOperatorCall(aggregate, letOperator,
                                                                     StringConstants.SelectMethod,
                                                                     ImmutableArray.Create(Of BoundExpression)(selectSelectorLambda),
                                                                     aggregate.AggregateKeyword.Span,
                                                                     callDiagnostics)
                End If

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

                result = New BoundAggregateClause(aggregate, Nothing, Nothing,
                                                  underlyingExpression,
                                                  firstSelectSelectorBinder.RangeVariables.Concat(declaredRangeVariables),
                                                  selectSelectorLambda.Expression.Type,
                                                  ImmutableArray.Create(Of Binder)(firstSelectSelectorBinder, intoBinder),
                                                  underlyingExpression.Type)
            End If

            Debug.Assert(Not result.Binders.IsDefault AndAlso
                             result.Binders.Length = 2 AndAlso
                             result.Binders(0) IsNot Nothing AndAlso
                             result.Binders(1) IsNot Nothing)

            Return result
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the outer, bind the following Group Join operator.
        ''' 
        '''     [{Preceding query operators}] Group Join {collection range variable} 
        '''                                              [{additional joins}] 
        '''                                   On {condition}
        '''                                   Into {aggregation range variables}
        ''' 
        ''' Ex: From a In AA Group Join b in BB          AA.GroupJoin(BB, Function(a) Key(a), Function(b) Key(b), 
        '''                  On Key(a) Equals Key(b) ==>              Function(a, group_b) New With {a, group_b.Count()})
        '''                  Into Count()
        ''' 
        ''' Also, depending on the amount of collection range variables in scope, and the following query operators,
        ''' translation can produce a nested, as opposed to flat, compound variable (see BindInnerJoinClause for an example).
        ''' 
        ''' Note, that type of the group must be inferred from the set of available GroupJoin operators in order to be able to 
        ''' interpret the aggregation range variables. 
        ''' </summary>
        Private Function BindGroupJoinClause(
                                              outer As BoundQueryClauseBase,
                                              groupJoin As GroupJoinClauseSyntax,
                                              declaredNames As PooledHashSet(Of String),
                                              operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                              diagnostics As DiagnosticBag
                                            ) As BoundQueryClause

            Debug.Assert((declaredNames IsNot Nothing) = (groupJoin.Parent.Kind.IsKindEither(SyntaxKind.SimpleJoinClause, SyntaxKind.GroupJoinClause)))

            ' Shadowing rules for range variables declared by [Group Join] are a little
            ' bit tricky:
            '  1) Range variables declared within the inner source (JoinedVariables + AdditionalJoins)
            '     can survive only up until the [On] clause, after that, even those in scope, move out of scope
            '     into the Group. Other range variables in scope in the [On] are the outer's range variables.
            '     Range variables from outer's outer are not in scope and, therefore, are never shadowed by the
            '     same-named inner source range variables. Range variables declared within the inner source that
            '     go out of scope before interpretation reaches the [On] clause do not shadow even same-named
            '     outer's range variables. 
            '
            '  2) Range variables declared in the [Into] clause must not shadow outer's range variables simply 
            '     because they are merged into the same Anonymous Type by the [Into] selector. They also must
            '     not shadow outer's outer range variables (possibly throughout the whole hierarchy),
            '     with which they will later get into the same scope within an [On] clause. Note, that declaredNames
            '     parameter, when passed, includes the names of all such range variables and this function will add
            '     to this set.

            Dim namesInScopeInOnClause = CreateSetOfDeclaredNames(outer.RangeVariables)

            If declaredNames Is Nothing Then
                Debug.Assert(groupJoin Is operatorsEnumerator.Current)
                declaredNames = CreateSetOfDeclaredNames(outer.RangeVariables)
            Else
                AssertDeclaredNames(declaredNames, outer.RangeVariables)
            End If

            Debug.Assert(groupJoin.JoinedVariables.Count = 1, "Malformed syntax tree.")

            Dim inner As BoundQueryClauseBase

            If groupJoin.JoinedVariables.Count = 0 Then
                ' Malformed tree.
                Dim rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, groupJoin, ErrorTypeSymbol.UnknownResultType)

                inner = New BoundQueryableSource(groupJoin,
                                                 New BoundQuerySource(BadExpression(groupJoin, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()).MakeCompilerGenerated(),
                                                 Nothing,
                                                 ImmutableArray.Create(rangeVar),
                                                 ErrorTypeSymbol.UnknownResultType,
                                                 ImmutableArray(Of Binder).Empty,
                                                 ErrorTypeSymbol.UnknownResultType,
                                                 hasErrors:=True).MakeCompilerGenerated()
            Else
                inner = BindCollectionRangeVariable(groupJoin.JoinedVariables(0), False, namesInScopeInOnClause, diagnostics)
            End If

            For Each additionalJoin As JoinClauseSyntax In groupJoin.AdditionalJoins
                Select Case additionalJoin.Kind
                    Case SyntaxKind.SimpleJoinClause
                        inner = BindInnerJoinClause(inner, DirectCast(additionalJoin, SimpleJoinClauseSyntax), namesInScopeInOnClause, Nothing, diagnostics)
                    Case SyntaxKind.GroupJoinClause
                        inner = BindGroupJoinClause(inner, DirectCast(additionalJoin, GroupJoinClauseSyntax), namesInScopeInOnClause, Nothing, diagnostics)
                End Select
            Next

            Debug.Assert(outer.RangeVariables.Length > 0 AndAlso inner.RangeVariables.Length > 0)
            AssertDeclaredNames(namesInScopeInOnClause, inner.RangeVariables)

            ' Bind keys.
            Dim outerKeyLambda As BoundQueryLambda = Nothing
            Dim innerKeyLambda As BoundQueryLambda = Nothing
            Dim outerKeyBinder As QueryLambdaBinder = Nothing
            Dim innerKeyBinder As QueryLambdaBinder = Nothing

            QueryLambdaBinder.BindJoinKeys(Me, groupJoin, outer, inner,
                                           outer.RangeVariables.Concat(inner.RangeVariables),
                                           outerKeyLambda, outerKeyBinder,
                                           innerKeyLambda, innerKeyBinder,
                                           diagnostics)

            ' Infer type of the resulting group.
            Dim methodGroup As BoundMethodGroup = Nothing
            Dim groupType As TypeSymbol = InferGroupType(outer, inner, groupJoin, outerKeyLambda, innerKeyLambda, methodGroup, diagnostics)

            ' Bind the INTO selector.
            Dim intoBinder As IntoClauseBinder = Nothing
            Dim intoRangeVariables As ImmutableArray(Of RangeVariableSymbol) = Nothing
            Dim intoLambda As BoundQueryLambda = BindIntoSelectorLambda(groupJoin, outer.RangeVariables, outer.CompoundVariableType,
                                                                        True, declaredNames,
                                                                        groupType, inner.RangeVariables, inner.CompoundVariableType,
                                                                        groupJoin.AggregationVariables,
                                                                        MustProduceFlatCompoundVariable(groupJoin, operatorsEnumerator),
                                                                        diagnostics, intoBinder, intoRangeVariables)

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If outer.Type.IsErrorType() OrElse methodGroup Is Nothing Then
                boundCallOrBadExpression = BadExpression(groupJoin, ImmutableArray.Create(Of BoundExpression)(outer, inner, outerKeyLambda, innerKeyLambda, intoLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim callDiagnostics As DiagnosticBag = diagnostics

                If inner.HasErrors OrElse inner.Type.IsErrorType() OrElse
                   ShouldSuppressDiagnostics(outerKeyLambda) OrElse
                   ShouldSuppressDiagnostics(innerKeyLambda) OrElse
                   ShouldSuppressDiagnostics(intoLambda) Then

                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    callDiagnostics = DiagnosticBag.GetInstance()
                End If

                ' Reusing method group that we got while inferring group type, this way we can avoid doing name lookup again. 
                boundCallOrBadExpression = BindQueryOperatorCall(groupJoin, outer,
                                                               StringConstants.GroupJoinMethod,
                                                               methodGroup,
                                                               ImmutableArray.Create(Of BoundExpression)(inner, outerKeyLambda, innerKeyLambda, intoLambda),
                                                               groupJoin.JoinKeyword.Span,
                                                               callDiagnostics)

                If callDiagnostics IsNot diagnostics Then
                    callDiagnostics.Free()
                End If
            End If
            namesInScopeInOnClause?.Free()
            'declaredNames?.Free()
            Return New BoundQueryClause(groupJoin,
                                        boundCallOrBadExpression,
                                        outer.RangeVariables.Concat(intoRangeVariables),
                                        intoLambda.Expression.Type,
                                        ImmutableArray.Create(Of Binder)(outerKeyBinder, innerKeyBinder, intoBinder),
                                        boundCallOrBadExpression.Type)

        End Function

        Private Shared Sub GetAbsorbingJoinSelectorLambdaKindAndSyntax(
                                                                        clauseSyntax As QueryClauseSyntax,
                                                                        absorbNextOperator As QueryClauseSyntax,
                                                            <Out> ByRef lambda As (Kind As SynthesizedLambdaKind,Syntax As VisualBasicSyntaxNode)
                                                                      )
            ' Join selector is either associated with absorbed select/let/aggregate clause
            ' or it doesn't contain user code (just pairs outer with inner into an anonymous type or is an identity).
            If absorbNextOperator Is Nothing Then
                Select Case clauseSyntax.Kind
                    Case SyntaxKind.SimpleJoinClause
                        lambda.Kind = SynthesizedLambdaKind.JoinNonUserCodeQueryLambda

                    Case SyntaxKind.FromClause
                        lambda.Kind = SynthesizedLambdaKind.FromNonUserCodeQueryLambda

                    Case SyntaxKind.AggregateClause
                        lambda.Kind = SynthesizedLambdaKind.FromNonUserCodeQueryLambda
                    Case SyntaxKind.ZipClause
                        lambda.Kind = SynthesizedLambdaKind.ZipNonUserCodeQueryLambda

                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(clauseSyntax.Kind)
                End Select

                lambda.Syntax = clauseSyntax
                Debug.Assert(LambdaUtilities.IsNonUserCodeQueryLambda(lambda.Syntax))
            Else
                Select Case absorbNextOperator.Kind
                    Case SyntaxKind.AggregateClause
                        lambda = (  Kind:=SynthesizedLambdaKind.AggregateQueryLambda,
                                  Syntax:=LambdaUtilities.GetFromOrAggregateVariableLambdaBody(DirectCast(absorbNextOperator, AggregateClauseSyntax).Variables.First))

                    Case SyntaxKind.LetClause
                        lambda = (  Kind:=SynthesizedLambdaKind.LetVariableQueryLambda,
                                  Syntax:=LambdaUtilities.GetLetVariableLambdaBody(DirectCast(absorbNextOperator, LetClauseSyntax).Variables.First))

                    Case SyntaxKind.SelectClause
                        lambda = (  Kind:=SynthesizedLambdaKind.SelectQueryLambda,
                                  Syntax:=LambdaUtilities.GetSelectLambdaBody(DirectCast(absorbNextOperator, SelectClauseSyntax)))

                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(absorbNextOperator.Kind)
                End Select

                Debug.Assert(LambdaUtilities.IsLambdaBody(lambda.Syntax))
            End If
        End Sub

        Private Shared Function JoinShouldAbsorbNextOperator(
                                                        ByRef operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator
                                                            ) As QueryClauseSyntax

            Dim copyOfOperatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator = operatorsEnumerator

            If Not copyOfOperatorsEnumerator.MoveNext() Then Return Nothing
            Dim nextOperator As QueryClauseSyntax = copyOfOperatorsEnumerator.Current

            Select Case nextOperator.Kind
                   Case SyntaxKind.LetClause
                        If DirectCast(nextOperator, LetClauseSyntax).Variables.Count > 0 Then
                           ' Absorb Let
                           operatorsEnumerator = copyOfOperatorsEnumerator
                           Return nextOperator
                        Else
                           ' Malformed tree.
                           Debug.Assert(DirectCast(nextOperator, LetClauseSyntax).Variables.Count > 0, "Malformed syntax tree.")
                        End If

                   Case SyntaxKind.SelectClause
                        ' Absorb Select
                        operatorsEnumerator = copyOfOperatorsEnumerator
                        Return nextOperator

                   Case SyntaxKind.AggregateClause
                        ' Absorb Aggregate
                        operatorsEnumerator = copyOfOperatorsEnumerator
                        Return nextOperator
            End Select

            Return Nothing
        End Function

        Private Function AbsorbOperatorFollowingJoin(
                                                      absorbingJoin As BoundQueryClause,
                                                      absorbNextOperator As QueryClauseSyntax,
                                                      operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                      joinSelectorDeclaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                      joinSelectorBinder As QueryLambdaBinder,
                                                      leftRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                      rightRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                      group As BoundQueryClauseBase,
                                                      intoBinder As IntoClauseDisallowGroupReferenceBinder,
                                                      diagnostics As DiagnosticBag
                                                    ) As BoundQueryClauseBase

            Debug.Assert(absorbNextOperator Is operatorsEnumerator.Current)
            Debug.Assert(absorbingJoin.Binders.Length > 1)

            Select Case absorbNextOperator.Kind
                   Case SyntaxKind.SelectClause
                       ' Absorb Select.
                       Return New BoundQueryClause(absorbNextOperator,
                                                   absorbingJoin, joinSelectorDeclaredRangeVariables,
                                                   absorbingJoin.CompoundVariableType,
                                                   ImmutableArray.Create(absorbingJoin.Binders.Last),
                                                   absorbingJoin.Type)

                   Case SyntaxKind.LetClause
                       ' Absorb Let.
                       ' First expression range variable was handled by the join selector,
                       ' create node for it.
                       Dim [let] = DirectCast(absorbNextOperator, LetClauseSyntax)
                       Debug.Assert([let].Variables.Count > 0)
                       Dim firstVariable As ExpressionRangeVariableSyntax = [let].Variables.First
                       Dim absorbedLet As New BoundQueryClause(firstVariable,
                                                               absorbingJoin,
                                                               absorbingJoin.RangeVariables.Concat(joinSelectorDeclaredRangeVariables),
                                                               absorbingJoin.CompoundVariableType,
                                                               ImmutableArray.Create(absorbingJoin.Binders.Last),
                                                               absorbingJoin.Type)

                       ' Handle the rest of the variables.
                       Return BindLetClause(absorbedLet, [let], operatorsEnumerator, diagnostics, skipFirstVariable:=True)

                   Case SyntaxKind.AggregateClause
                       ' Absorb Aggregate.

                       Return CompleteAggregateClauseBinding(DirectCast(absorbNextOperator, AggregateClauseSyntax),
                                                             operatorsEnumerator,
                                                             leftRangeVariables,
                                                             rightRangeVariables,
                                                             absorbingJoin,
                                                             joinSelectorBinder,
                                                             joinSelectorDeclaredRangeVariables,
                                                             absorbingJoin.CompoundVariableType,
                                                             group,
                                                             intoBinder,
                                                             diagnostics)
            End Select
            Throw ExceptionUtilities.UnexpectedValue(absorbNextOperator.Kind)
        End Function

        ''' <summary>
        ''' Given result of binding preceding query operators, the outer, bind the following Join operator.
        ''' 
        '''     [{Preceding query operators}] Join {collection range variable} 
        '''                                        [{additional joins}] 
        '''                                   On {condition}
        ''' 
        ''' Ex: From a In AA Join b in BB On Key(a) Equals Key(b)  ==> AA.Join(BB, Function(a) Key(a), Function(b) Key(b), 
        '''                                                                    Function(a, b) New With {a, b})
        ''' 
        ''' Ex: From a In AA                       AA.Join(
        '''     Join b in BB                               BB.Join(CC, Function(b) Key(b), Function(c) Key(c),
        '''          Join c in CC             ==>                  Function(b, c) New With {b, c}),
        '''          On Key(c) Equals Key(b)               Function(a) Key(a), Function({b, c}) Key(b),
        '''     On Key(a) Equals Key(b)                    Function(a, {b, c}) New With {a, b, c})
        '''                                                                    
        ''' 
        ''' Also, depending on the amount of collection range variables in scope, and the following query operators,
        ''' translation can produce a nested, as opposed to flat, compound variable.
        ''' 
        ''' Ex: From a In AA                       AA.Join(BB, Function(a) Key(a), Function(b) Key(b),
        '''     Join b in BB                               Function(a, b) New With {a, b}).
        '''     On Key(a) Equals Key(b)               Join(CC, Function({a, b}) Key(a, b), Function(c) Key(c),
        '''     Join c in CC             ==>               Function({a, b}, c) New With {{a, b}, c}).
        '''     On Key(c) Equals Key(a, b)            Join(DD, Function({{a, b}, c}) Key(a, b, c), Function(d) Key(d),
        '''     Join d in DD                               Function({{a, b}, c}, d) New With {a, b, c, d})
        '''     On Key(a, b, c) Equals Key(d)
        ''' 
        ''' If Join is immediately followed by a Select or a Let operator, they are absorbed by the translation. 
        ''' When this happens, operatorsEnumerator is advanced appropriately.
        ''' 
        ''' Ex: From a In AA Join b in BB On Key(a) Equals Key(b)  ==> AA.Join(BB, Function(a) Key(a), Function(b) Key(b), 
        '''     Select a + b                                                   Function(a, b) a + b)
        ''' 
        ''' Ex: From a In AA Join b in BB On Key(a) Equals Key(b)  ==> AA.Join(BB, Function(a) Key(a), Function(b) Key(b), 
        '''     Let c                                                   Function(a, b) New With {a, b, c})
        ''' 
        ''' </summary>
        Private Function BindInnerJoinClause(
                                              outer As BoundQueryClauseBase,
                                              join As SimpleJoinClauseSyntax,
                                              declaredNames As PooledHashSet(Of String),
                                        ByRef operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                              diagnostics As DiagnosticBag
                                            ) As BoundQueryClauseBase

            Debug.Assert(join.Kind = SyntaxKind.SimpleJoinClause)
            Debug.Assert((declaredNames IsNot Nothing) = (join.Parent.Kind.IsKindEither(SyntaxKind.SimpleJoinClause, SyntaxKind.GroupJoinClause)))

            Dim isNested As Boolean

            If declaredNames Is Nothing Then
                Debug.Assert(join Is operatorsEnumerator.Current)
                isNested = False
                declaredNames = CreateSetOfDeclaredNames(outer.RangeVariables)
            Else
                isNested = True
                AssertDeclaredNames(declaredNames, outer.RangeVariables)
            End If

            Debug.Assert(join.JoinedVariables.Count = 1, "Malformed syntax tree.")

            Dim inner As BoundQueryClauseBase

            If join.JoinedVariables.Count = 0 Then
                ' Malformed tree.
                Dim rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, join, ErrorTypeSymbol.UnknownResultType)

                inner = New BoundQueryableSource(join,
                                                 New BoundQuerySource(BadExpression(join, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()).MakeCompilerGenerated(),
                                                 Nothing,
                                                 ImmutableArray.Create(rangeVar),
                                                 ErrorTypeSymbol.UnknownResultType,
                                                 ImmutableArray(Of Binder).Empty,
                                                 ErrorTypeSymbol.UnknownResultType,
                                                 hasErrors:=True).MakeCompilerGenerated()
            Else
                inner = BindCollectionRangeVariable(join.JoinedVariables(0), False, declaredNames, diagnostics)
            End If


            For Each additionalJoin As JoinClauseSyntax In join.AdditionalJoins
                Select Case additionalJoin.Kind
                    Case SyntaxKind.SimpleJoinClause
                        inner = BindInnerJoinClause(inner, DirectCast(additionalJoin, SimpleJoinClauseSyntax), declaredNames, Nothing, diagnostics)
                    Case SyntaxKind.GroupJoinClause
                        inner = BindGroupJoinClause(inner, DirectCast(additionalJoin, GroupJoinClauseSyntax), declaredNames, Nothing, diagnostics)
                End Select
            Next

            AssertDeclaredNames(declaredNames, inner.RangeVariables)

            ' Bind keys.
            Dim outerKeyLambda As BoundQueryLambda = Nothing
            Dim innerKeyLambda As BoundQueryLambda = Nothing
            Dim outerKeyBinder As QueryLambdaBinder = Nothing
            Dim innerKeyBinder As QueryLambdaBinder = Nothing
            Dim joinSelectorRangeVariables As ImmutableArray(Of RangeVariableSymbol) = outer.RangeVariables.Concat(inner.RangeVariables)

            QueryLambdaBinder.BindJoinKeys(Me, join, outer, inner,
                                           joinSelectorRangeVariables,
                                           outerKeyLambda, outerKeyBinder,
                                           innerKeyLambda, innerKeyBinder,
                                           diagnostics)

            ' Create LambdaSymbol for the shape of the join-selector lambda.
            Dim joinSelectorParamLeft As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterNameLeft(outer.RangeVariables), 0,
                                                                                                       outer.CompoundVariableType,
                                                                                                       join, outer.RangeVariables)

            Dim joinSelectorParamRight As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterNameRight(inner.RangeVariables), 1,
                                                                                                        inner.CompoundVariableType,
                                                                                                        join, inner.RangeVariables)

            Dim lambdaBinders As ImmutableArray(Of Binder)

            ' If the next operator is a Select or a Let, we should absorb it by putting its selector
            ' in the join lambda.
            Dim absorbNextOperator As QueryClauseSyntax = Nothing

            If Not isNested Then absorbNextOperator = JoinShouldAbsorbNextOperator(operatorsEnumerator)

            Dim joinSelectorDeclaredRangeVariables As ImmutableArray(Of RangeVariableSymbol)
            Dim joinSelector As BoundExpression
            Dim group As BoundQueryClauseBase = Nothing
            Dim intoBinder As IntoClauseDisallowGroupReferenceBinder = Nothing

            Dim joinSelectorLambda As (Kind As SynthesizedLambdaKind, Syntax As VisualBasicSyntaxNode)  = Nothing

            GetAbsorbingJoinSelectorLambdaKindAndSyntax(join, absorbNextOperator, joinSelectorLambda)

            Dim joinSelectorLambdaSymbol = Me.CreateQueryLambdaSymbol(joinSelectorLambda,
                                                                      ImmutableArray.Create(joinSelectorParamLeft, joinSelectorParamRight))

            Dim joinSelectorBinder As New QueryLambdaBinder(joinSelectorLambdaSymbol, joinSelectorRangeVariables)

            If absorbNextOperator IsNot Nothing Then

                ' Absorb selector of the next operator.
                joinSelectorDeclaredRangeVariables = Nothing
                joinSelectorLambda.Syntax = Nothing
                joinSelector = joinSelectorBinder.BindAbsorbingJoinSelector(absorbNextOperator,
                                                                            operatorsEnumerator,
                                                                            outer.RangeVariables,
                                                                            inner.RangeVariables,
                                                                            joinSelectorDeclaredRangeVariables,
                                                                            group,
                                                                            intoBinder,
                                                                            diagnostics)

                lambdaBinders = ImmutableArray.Create(Of Binder)(outerKeyBinder, innerKeyBinder, joinSelectorBinder)
            Else
                Debug.Assert(outer.RangeVariables.Length > 0 AndAlso inner.RangeVariables.Length > 0)
                joinSelectorDeclaredRangeVariables = ImmutableArray(Of RangeVariableSymbol).Empty

                ' Need to build an Anonymous Type.

                ' In some scenarios, it is safe to leave compound variable in nested form when there is an
                ' operator down the road that does its own projection (Select, Group By, ...). 
                ' All following operators have to take an Anonymous Type in both cases and, since there is no way to
                ' restrict the shape of the Anonymous Type in method's declaration, the operators should be
                ' insensitive to the shape of the Anonymous Type.
                joinSelector = joinSelectorBinder.BuildJoinSelector(join,
                                                                    MustProduceFlatCompoundVariable(join, operatorsEnumerator),
                                                                    diagnostics)

                ' Not including joinSelectorBinder because there is no syntax behind this joinSelector,
                ' it is purely synthetic.
                lambdaBinders = ImmutableArray.Create(Of Binder)(outerKeyBinder, innerKeyBinder)
            End If

            Dim boundJoinQueryLambda = CreateBoundQueryLambda(joinSelectorLambdaSymbol,
                                                              joinSelectorRangeVariables,
                                                              joinSelector,
                                                              exprIsOperandOfConditionalBranch:=False)

            joinSelectorLambdaSymbol.SetQueryLambdaReturnType(joinSelector.Type)
            boundJoinQueryLambda.SetWasCompilerGenerated()

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If outer.Type.IsErrorType() Then
                boundCallOrBadExpression = BadExpression(join, ImmutableArray.Create(Of BoundExpression)(outer, inner, outerKeyLambda, innerKeyLambda, boundJoinQueryLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim callDiagnostics As DiagnosticBag = diagnostics
                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If inner.HasErrors OrElse inner.Type.IsErrorType() OrElse
                   ShouldSuppressDiagnostics(outerKeyLambda) OrElse
                   ShouldSuppressDiagnostics(innerKeyLambda) OrElse
                   ShouldSuppressDiagnostics(boundJoinQueryLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    callDiagnostics = suppressDiagnostics
                End If

                boundCallOrBadExpression = BindQueryOperatorCall(join, outer,
                                                                 StringConstants.JoinMethod,
                                                                 ImmutableArray.Create(Of BoundExpression)(inner, outerKeyLambda, innerKeyLambda, boundJoinQueryLambda),
                                                                 join.JoinKeyword.Span,
                                                                 callDiagnostics)

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            End If

            Dim result As BoundQueryClauseBase = New BoundQueryClause(join,
                                                                      boundCallOrBadExpression,
                                                                      joinSelectorRangeVariables,
                                                                      boundJoinQueryLambda.Expression.Type,
                                                                      lambdaBinders,
                                                                      boundCallOrBadExpression.Type)
           ' declaredNames?.Free
            If absorbNextOperator Is Nothing Then Return result

            Debug.Assert(Not isNested)
            Return AbsorbOperatorFollowingJoin(DirectCast(result, BoundQueryClause),
                                               absorbNextOperator,
                                               operatorsEnumerator,
                                               joinSelectorDeclaredRangeVariables,
                                               joinSelectorBinder,
                                               outer.RangeVariables,
                                               inner.RangeVariables,
                                               group,
                                               intoBinder,
                                               diagnostics)
        End Function

    End Class

End Namespace

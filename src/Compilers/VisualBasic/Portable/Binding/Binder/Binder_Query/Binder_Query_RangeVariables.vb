' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Shared Function ShadowsRangeVariableInTheChildScope(
                                                                     childScopeBinder As Binder,
                                                                     rangeVar As RangeVariableSymbol
                                                                   ) As Boolean
            Dim lookup = LookupResult.GetInstance()

            childScopeBinder.LookupInSingleBinder(lookup, rangeVar.Name, 0, Nothing, childScopeBinder, useSiteDiagnostics:=Nothing)

            Dim result As Boolean = (lookup.IsGood AndAlso lookup.Symbols(0).Kind = SymbolKind.RangeVariable)

            lookup.Free()

            Return result
        End Function

        ''' <summary>
        ''' See comments for BindFromClause method, this method actually does all the work.
        ''' </summary>
        Private Function BindCollectionRangeVariables(
                                                       clauseSyntax As QueryClauseSyntax,
                                                       sourceOpt As BoundQueryClauseBase,
                                                       variables As SeparatedSyntaxList(Of CollectionRangeVariableSyntax),
                                                 ByRef operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                       diagnostics As DiagnosticBag
                                                     ) As BoundQueryClauseBase

            Debug.Assert(clauseSyntax.Kind.IsKindEither(SyntaxKind.AggregateClause,
                                                    SyntaxKind.FromClause,
                                                    SyntaxKind.ZipClause))
            Debug.Assert(variables.Count > 0, "Malformed syntax tree.")

            ' Handle malformed tree gracefully.
            If variables.Count = 0 Then Return BindEmptyCollectionRangeVariables(clauseSyntax, sourceOpt)

            Dim source As BoundQueryClauseBase = sourceOpt

            If source Is Nothing Then
                ' We are at the beginning of the query.
                ' Let's go ahead and process the first collection range variable then.
                source = BindCollectionRangeVariable(variables(0), True, Nothing, diagnostics)
                Debug.Assert(source.RangeVariables.Length = 1)
            End If

            Dim suppressDiagnostics As DiagnosticBag = Nothing
            Dim callDiagnostics As DiagnosticBag = diagnostics

            For i = If(source Is sourceOpt, 0, 1) To variables.Count - 1

                Dim variable As CollectionRangeVariableSyntax = variables(i)

                ' Create LambdaSymbol for the shape of the many-selector lambda.
                Dim manySelectorParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                                       source.CompoundVariableType,
                                                                                                       variable, source.RangeVariables)

                Dim manySelectorLambdaSymbol = Me.CreateQueryLambdaSymbol(
                                                                          (SynthesizedLambdaKind.FromOrAggregateVariableQueryLambda,LambdaUtilities.GetFromOrAggregateVariableLambdaBody(variable)),
                                                                          ImmutableArray.Create(manySelectorParam))

                ' Create binder for the many selector.
                Dim manySelectorBinder As New QueryLambdaBinder(manySelectorLambdaSymbol, source.RangeVariables)

                Dim manySelector As BoundQueryableSource = manySelectorBinder.BindCollectionRangeVariable(variable, False, Nothing, diagnostics)
                Debug.Assert(manySelector.RangeVariables.Length = 1)

                Dim manySelectorLambda = CreateBoundQueryLambda(manySelectorLambdaSymbol,
                                                                source.RangeVariables,
                                                                manySelector,
                                                                exprIsOperandOfConditionalBranch:=False)

                ' Note, we are not setting return type for the manySelectorLambdaSymbol because
                ' we want it to be taken from the target delegate type. We don't care what it is going to be
                ' because it doesn't affect types of range variables after this operator.
                manySelectorLambda.SetWasCompilerGenerated()

                ' Create LambdaSymbol for the shape of the join-selector lambda.
                Dim joinSelectorParamLeftSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterNameLeft(source.RangeVariables), 0,
                                                                                                           source.CompoundVariableType,
                                                                                                           variable, source.RangeVariables)

                Dim joinSelectorParamRightSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterNameRight(manySelector.RangeVariables), 1,
                                                                                                            manySelector.CompoundVariableType,
                                                                                                            variable, manySelector.RangeVariables)
                Dim lambdaBinders As ImmutableArray(Of Binder)

                ' If this is the last collection range variable, see if the next operator is
                ' a Select or a Let. If it is, we should absorb it by putting its selector
                ' in the join lambda.
                Dim absorbNextOperator As QueryClauseSyntax = Nothing

                If i = variables.Count - 1 Then absorbNextOperator = JoinShouldAbsorbNextOperator(operatorsEnumerator)

                Dim sourceRangeVariables = source.RangeVariables
                Dim joinSelectorRangeVariables As ImmutableArray(Of RangeVariableSymbol) = sourceRangeVariables.Concat(manySelector.RangeVariables)
                Dim joinSelectorDeclaredRangeVariables As ImmutableArray(Of RangeVariableSymbol)
                Dim joinSelector As BoundExpression
                Dim group As BoundQueryClauseBase = Nothing
                Dim intoBinder As IntoClauseDisallowGroupReferenceBinder = Nothing
                Dim joinSelectorBinder As QueryLambdaBinder = Nothing

                Dim joinSelectorLambda As (Kind As SynthesizedLambdaKind, Syntax As VisualBasicSyntaxNode) = Nothing
                GetAbsorbingJoinSelectorLambdaKindAndSyntax(clauseSyntax, absorbNextOperator, joinSelectorLambda)

                Dim joinSelectorLambdaSymbol = Me.CreateQueryLambdaSymbol(joinSelectorLambda, ImmutableArray.Create(joinSelectorParamLeftSymbol, joinSelectorParamRightSymbol))

                If absorbNextOperator IsNot Nothing Then

                    ' Absorb selector of the next operator.
                    joinSelectorBinder = New QueryLambdaBinder(joinSelectorLambdaSymbol, joinSelectorRangeVariables)

                    joinSelectorDeclaredRangeVariables = Nothing
                    joinSelector = joinSelectorBinder.BindAbsorbingJoinSelector(absorbNextOperator,
                                                                                operatorsEnumerator,
                                                                                sourceRangeVariables,
                                                                                manySelector.RangeVariables,
                                                                                joinSelectorDeclaredRangeVariables,
                                                                                group,
                                                                                intoBinder,
                                                                                diagnostics)

                    lambdaBinders = ImmutableArray.Create(Of Binder)(manySelectorBinder, joinSelectorBinder)
                Else
                    joinSelectorDeclaredRangeVariables = ImmutableArray(Of RangeVariableSymbol).Empty

                    If sourceRangeVariables.Length > 0 Then
                        ' Need to build an Anonymous Type.
                        joinSelectorBinder = New QueryLambdaBinder(joinSelectorLambdaSymbol, joinSelectorRangeVariables)

                        ' If it is not the last variable in the list, we simply combine source's
                        ' compound variable (an instance of its Anonymous Type) with our new variable, 
                        ' creating new compound variable of nested Anonymous Type.

                        ' In some scenarios, it is safe to leave compound variable in nested form when there is an
                        ' operator down the road that does its own projection (Select, Group By, ...). 
                        ' All following operators have to take an Anonymous Type in both cases and, since there is no way to
                        ' restrict the shape of the Anonymous Type in method's declaration, the operators should be
                        ' insensitive to the shape of the Anonymous Type.
                        joinSelector = joinSelectorBinder.BuildJoinSelector(variable,
                                                                            (i = variables.Count - 1 AndAlso
                                                                                MustProduceFlatCompoundVariable(operatorsEnumerator)),
                                                                            diagnostics)
                    Else
                        ' Easy case, no need to build an Anonymous Type.
                        Debug.Assert(sourceRangeVariables.Length = 0)
                        joinSelector = New BoundParameter(joinSelectorParamRightSymbol.Syntax, joinSelectorParamRightSymbol,
                                                           False, joinSelectorParamRightSymbol.Type).MakeCompilerGenerated()
                    End If

                    lambdaBinders = ImmutableArray.Create(Of Binder)(manySelectorBinder)
                End If

                ' Join selector is either associated with absorbed select/let/aggregate clause
                ' or it doesn't contain user code (just pairs outer with inner into an anonymous type or is an identity).
                Dim boundJoinQuetyLambda = CreateBoundQueryLambda(joinSelectorLambdaSymbol,
                                                                joinSelectorRangeVariables,
                                                                joinSelector,
                                                                exprIsOperandOfConditionalBranch:=False)

                joinSelectorLambdaSymbol.SetQueryLambdaReturnType(joinSelector.Type)
                boundJoinQuetyLambda.SetWasCompilerGenerated()

                ' Now bind the call.
                Dim boundCallOrBadExpression As BoundExpression

                If source.Type.IsErrorType() Then
                    boundCallOrBadExpression = BadExpression(variable, ImmutableArray.Create(Of BoundExpression)(source, manySelectorLambda, boundJoinQuetyLambda),
                                                             ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                Else
                    If suppressDiagnostics Is Nothing AndAlso
                       (ShouldSuppressDiagnostics(manySelectorLambda) OrElse ShouldSuppressDiagnostics(boundJoinQuetyLambda)) Then
                        ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                        suppressDiagnostics = DiagnosticBag.GetInstance()
                        callDiagnostics = suppressDiagnostics
                    End If

                    Dim operatorNameLocation As TextSpan

                    If i = 0 Then
                        ' This is the first variable.
                        operatorNameLocation = clauseSyntax.GetFirstToken().Span
                    Else
                        operatorNameLocation = variables.GetSeparator(i - 1).Span
                    End If

                    boundCallOrBadExpression = BindQueryOperatorCall(variable, source,
                                                                     StringConstants.SelectManyMethod,
                                                                     ImmutableArray.Create(Of BoundExpression)(manySelectorLambda, boundJoinQuetyLambda),
                                                                     operatorNameLocation,
                                                                     callDiagnostics)
                End If

                source = New BoundQueryClause(variable,
                                              boundCallOrBadExpression,
                                              joinSelectorRangeVariables,
                                              boundJoinQuetyLambda.Expression.Type,
                                              lambdaBinders,
                                              boundCallOrBadExpression.Type)

                If absorbNextOperator IsNot Nothing Then
                    Debug.Assert(i = variables.Count - 1)
                    source = AbsorbOperatorFollowingJoin(DirectCast(source, BoundQueryClause),
                                                         absorbNextOperator, operatorsEnumerator,
                                                         joinSelectorDeclaredRangeVariables,
                                                         joinSelectorBinder,
                                                         sourceRangeVariables,
                                                         manySelector.RangeVariables,
                                                         group,
                                                         intoBinder,
                                                         diagnostics)
                    Exit For
                End If
            Next

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Return source
        End Function

        Private Function BindEmptyCollectionRangeVariables(
                                                            clauseSyntax As QueryClauseSyntax,
                                                            sourceOpt As BoundQueryClauseBase
                                                          ) As BoundQueryClauseBase

            Dim rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, clauseSyntax, ErrorTypeSymbol.UnknownResultType)

            If sourceOpt Is Nothing Then
                Return New BoundQueryableSource(clauseSyntax,
                                                New BoundQuerySource(BadExpression(clauseSyntax, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()).MakeCompilerGenerated(),
                                                Nothing,
                                                ImmutableArray.Create(rangeVar),
                                                ErrorTypeSymbol.UnknownResultType,
                                                ImmutableArray.Create(Me),
                                                ErrorTypeSymbol.UnknownResultType,
                                                hasErrors:=True)
            Else
                Return New BoundQueryClause(clauseSyntax,
                                            BadExpression(clauseSyntax, sourceOpt, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated(),
                                            sourceOpt.RangeVariables.Add(rangeVar),
                                            ErrorTypeSymbol.UnknownResultType,
                                            ImmutableArray.Create(Me),
                                            ErrorTypeSymbol.UnknownResultType,
                                            hasErrors:=True)
            End If
        End Function

        Private Sub VerifyRangeVariableName(
                                             rangeVar As RangeVariableSymbol,
                                             identifier As SyntaxToken,
                                             diagnostics As DiagnosticBag
                                           )

            Debug.Assert(identifier.Parent Is rangeVar.Syntax)

            If identifier.GetTypeCharacter() <> TypeCharacter.None Then ReportDiagnostic(diagnostics, identifier, ERRID.ERR_QueryAnonymousTypeDisallowsTypeChar)

            If Compilation.ObjectType.GetMembers(rangeVar.Name).Length > 0 Then
                ReportDiagnostic(diagnostics, identifier, ERRID.ERR_QueryInvalidControlVariableName1)
            Else
                VerifyNameShadowingInMethodBody(rangeVar, identifier, identifier, diagnostics)
            End If
        End Sub

        Private Shared Function GetQueryOperatorNameSpan(
                                                    ByRef left As SyntaxToken,
                                                    ByRef right As SyntaxToken
                                                        ) As TextSpan

            Dim operatorNameSpan As TextSpan = left.Span

            If right.ValueText.Length > 0 Then operatorNameSpan = TextSpan.FromBounds(operatorNameSpan.Start, right.Span.End)

            Return operatorNameSpan
        End Function

        ''' <summary>
        ''' Bind CollectionRangeVariableSyntax, applying AsQueryable/AsEnumerable/Cast(Of Object) calls and 
        ''' Select with implicit type conversion as appropriate.
        ''' </summary>
        Private Function BindCollectionRangeVariable(
                                                      syntax As CollectionRangeVariableSyntax,
                                                      beginsTheQuery As Boolean,
                                                      declaredNames As PooledHashSet(Of String),
                                                      diagnostics As DiagnosticBag
                                                    ) As BoundQueryableSource

            Debug.Assert(declaredNames Is Nothing OrElse syntax.Parent.Kind.IsKindEither(SyntaxKind.SimpleJoinClause, SyntaxKind.GroupJoinClause))

            Dim source As BoundQueryPart = New BoundQuerySource(BindRValue(syntax.Expression, diagnostics))

            Dim variableType As TypeSymbol = Nothing
            Dim queryable As BoundExpression = ConvertToQueryableType(source, diagnostics, variableType)

            Dim sourceIsNotQueryable As Boolean = False

            If variableType Is Nothing Then
                If Not source.HasErrors Then
                    ReportDiagnostic(diagnostics, syntax.Expression, ERRID.ERR_ExpectedQueryableSource, source.Type)
                End If

                sourceIsNotQueryable = True

            ElseIf source IsNot queryable Then
                source = New BoundToQueryableCollectionConversion(DirectCast(queryable, BoundCall)).MakeCompilerGenerated()
            End If

            ' Deal with AsClauseOpt and various modifiers.
            Dim targetVariableType As TypeSymbol = Nothing

            If syntax.AsClause IsNot Nothing Then
                targetVariableType = DecodeModifiedIdentifierType(syntax.Identifier,
                                                                  syntax.AsClause,
                                                                  Nothing,
                                                                  Nothing,
                                                                  diagnostics,
                                                                  ModifiedIdentifierTypeDecoderContext.LocalType Or
                                                                  ModifiedIdentifierTypeDecoderContext.QueryRangeVariableType)
            ElseIf syntax.Identifier.Nullable.Node IsNot Nothing Then
                ReportDiagnostic(diagnostics, syntax.Identifier.Nullable, ERRID.ERR_NullableTypeInferenceNotSupported)
            End If

            If variableType Is Nothing Then
                Debug.Assert(sourceIsNotQueryable)

                If targetVariableType Is Nothing Then
                    variableType = ErrorTypeSymbol.UnknownResultType
                Else
                    variableType = targetVariableType
                End If

            ElseIf Not targetVariableType?.IsSameTypeIgnoringAll(variableType) Then
                Debug.Assert(Not sourceIsNotQueryable AndAlso syntax.AsClause IsNot Nothing)
                ' Need to apply implicit Select that converts variableType to targetVariableType.
                source = ApplyImplicitCollectionConversion(syntax, source, variableType, targetVariableType, diagnostics)
                variableType = targetVariableType
            End If

            Dim variable As RangeVariableSymbol = Nothing
            Dim rangeVarName As String = syntax.Identifier.Identifier.ValueText
            Dim rangeVariableOpt As RangeVariableSymbol = Nothing

            If rangeVarName?.Length = 0 Then
                ' Empty string must have been a syntax error. 
                rangeVarName = Nothing
            End If

            If rangeVarName IsNot Nothing Then
                variable = RangeVariableSymbol.Create(Me, syntax.Identifier.Identifier, variableType)

                ' Note what we are doing here:
                ' We are capturing rangeVariableOpt before doing any shadowing checks
                ' so that SemanticModel can find the declared symbol, but, if the variable will conflict with another
                ' variable in the same child scope, we will not add it to the scope. Instead, we create special
                ' error recovery range variable symbol and add it to the scope at the same place, making sure 
                ' that the earlier declared range variable wins during name lookup.
                ' As an alternative, we could still add the original range variable symbol to the scope and then,
                ' while we are binding the rest of the query in error recovery mode, references to the name would 
                ' cause ambiguity. However, this could negatively affect IDE experience. Also, as we build an
                ' Anonymous Type for the compound range variables, we would end up with a type with duplicate members,
                ' which could cause problems elsewhere.
                rangeVariableOpt = variable

                Dim doErrorRecovery As Boolean = False

                If declaredNames IsNot Nothing AndAlso Not declaredNames.Add(rangeVarName) Then
                    ReportDiagnostic(diagnostics, syntax.Identifier.Identifier, ERRID.ERR_QueryDuplicateAnonTypeMemberName1, rangeVarName)
                    doErrorRecovery = True  ' Shouldn't add to the scope.
                Else

                    ' Check shadowing etc. 
                    VerifyRangeVariableName(variable, syntax.Identifier.Identifier, diagnostics)

                    If Not beginsTheQuery AndAlso declaredNames Is Nothing Then
                        Debug.Assert(syntax.Parent.Kind.IsKindEither(SyntaxKind.FromClause,
                                                                     SyntaxKind.AggregateClause,
                                                                     SyntaxKind.ZipClause))
                        ' We are about to add this range variable to the current child scope.
                        If ShadowsRangeVariableInTheChildScope(Me, variable) Then
                            ' Shadowing error was reported earlier.
                            doErrorRecovery = True  ' Shouldn't add to the scope.
                        End If
                    End If
                End If

                If doErrorRecovery Then
                    variable = RangeVariableSymbol.CreateForErrorRecovery(Me, variable.Syntax, variableType)
                End If

            Else
                Debug.Assert(variable Is Nothing)
                variable = RangeVariableSymbol.CreateForErrorRecovery(Me, syntax, variableType)
            End If

            declaredNames.Free()
            Dim result As New BoundQueryableSource(syntax, source, rangeVariableOpt,
                                                   ImmutableArray.Create(variable), variableType,
                                                   ImmutableArray(Of Binder).Empty,
                                                   If(sourceIsNotQueryable, ErrorTypeSymbol.UnknownResultType, source.Type),
                                                   hasErrors:=sourceIsNotQueryable)

            Return result
        End Function

    End Class

End Namespace

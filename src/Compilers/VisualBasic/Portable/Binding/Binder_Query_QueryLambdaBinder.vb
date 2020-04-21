' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Collections
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' This is a top level binder used to bind bodies of query lambdas.
        ''' It also contains a bunch of helper methods to bind bodies of a particular kind.
        ''' </summary>
        Private NotInheritable Class QueryLambdaBinder
            Inherits BlockBaseBinder(Of RangeVariableSymbol) ' BlockBaseBinder implements various lookup methods as we need.

            Private ReadOnly _lambdaSymbol As LambdaSymbol
            Private ReadOnly _rangeVariables As ImmutableArray(Of RangeVariableSymbol)

            Public Sub New(lambdaSymbol As LambdaSymbol, rangeVariables As ImmutableArray(Of RangeVariableSymbol))
                MyBase.New(lambdaSymbol.ContainingBinder)
                _lambdaSymbol = lambdaSymbol
                _rangeVariables = rangeVariables
            End Sub

            Friend Overrides ReadOnly Property Locals As ImmutableArray(Of RangeVariableSymbol)
                Get
                    Return _rangeVariables
                End Get
            End Property

            Public ReadOnly Property RangeVariables As ImmutableArray(Of RangeVariableSymbol)
                Get
                    Return _rangeVariables
                End Get
            End Property

            Public Overrides ReadOnly Property ContainingMember As Symbol
                Get
                    Return _lambdaSymbol
                End Get
            End Property

            Public Overrides ReadOnly Property AdditionalContainingMembers As ImmutableArray(Of Symbol)
                Get
                    Return ImmutableArray(Of Symbol).Empty
                End Get
            End Property

            Public ReadOnly Property LambdaSymbol As LambdaSymbol
                Get
                    Return _lambdaSymbol
                End Get
            End Property

            Public Overrides ReadOnly Property IsInQuery As Boolean
                Get
                    Return True
                End Get
            End Property

            ''' <summary>
            ''' Bind body of a lambda representing Select operator selector in context of this binder.
            ''' </summary>
            Public Function BindSelectClauseSelector(
                [select] As SelectClauseSyntax,
                operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                <Out()> ByRef declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(declaredRangeVariables.IsDefault)

                Dim selectVariables As SeparatedSyntaxList(Of ExpressionRangeVariableSyntax) = [select].Variables
                Dim requireRangeVariable As Boolean

                If selectVariables.Count = 1 Then
                    requireRangeVariable = False

                    ' If there is a Join or a GroupJoin ahead, this operator must
                    ' add a range variable into the scope.
                    While operatorsEnumerator.MoveNext()
                        Select Case operatorsEnumerator.Current.Kind
                            Case SyntaxKind.SimpleJoinClause, SyntaxKind.GroupJoinClause
                                requireRangeVariable = True
                                Exit While

                            Case SyntaxKind.SelectClause, SyntaxKind.LetClause, SyntaxKind.FromClause
                                Exit While

                            Case SyntaxKind.AggregateClause
                                Exit While

                            Case SyntaxKind.GroupByClause
                                Exit While
                        End Select
                    End While
                Else
                    requireRangeVariable = True
                End If

                Return BindExpressionRangeVariables(selectVariables,
                                                    requireRangeVariable,
                                                    [select],
                                                    declaredRangeVariables,
                                                    diagnostics)
            End Function

            ''' <summary>
            ''' Bind Select like selector based on the set of expression range variables in context of this binder.
            ''' </summary>
            Public Function BindExpressionRangeVariables(
                selectVariables As SeparatedSyntaxList(Of ExpressionRangeVariableSyntax),
                requireRangeVariable As Boolean,
                selectorSyntax As QueryClauseSyntax,
                <Out()> ByRef declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(selectorSyntax IsNot Nothing AndAlso declaredRangeVariables.IsDefault)

                Dim selector As BoundExpression

                If selectVariables.Count = 1 Then
                    ' Simple case, one item in the projection.
                    Dim item As ExpressionRangeVariableSyntax = selectVariables(0)
                    selector = Nothing ' Suppress definite assignment error.

                    Debug.Assert(item.NameEquals Is Nothing OrElse item.NameEquals.AsClause Is Nothing)

                    ' Using enclosing binder for shadowing check, because range variables brought in scope by the current binder
                    ' are no longer in scope after the Select. So, it is fine to shadow them. 
                    Dim rangeVar As RangeVariableSymbol = Me.BindExpressionRangeVariable(item, requireRangeVariable, Me.ContainingBinder, Nothing, selector, diagnostics)

                    If rangeVar IsNot Nothing Then
                        declaredRangeVariables = ImmutableArray.Create(Of RangeVariableSymbol)(rangeVar)
                    Else
                        Debug.Assert(Not requireRangeVariable)
                        declaredRangeVariables = ImmutableArray(Of RangeVariableSymbol).Empty
                    End If

                ElseIf selectVariables.Count = 0 Then
                    ' Malformed tree
                    Debug.Assert(selectVariables.Count > 0, "Malformed syntax tree.")
                    selector = BadExpression(selectorSyntax, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                    declaredRangeVariables = ImmutableArray.Create(RangeVariableSymbol.CreateForErrorRecovery(Me, selectorSyntax, selector.Type))
                Else
                    ' Need to build an Anonymous Type
                    Dim selectors = New BoundExpression(selectVariables.Count - 1) {}
                    Dim declaredNames As HashSet(Of String) = CreateSetOfDeclaredNames()
                    Dim rangeVariables(selectVariables.Count - 1) As RangeVariableSymbol

                    For i As Integer = 0 To selectVariables.Count - 1
                        Debug.Assert(selectVariables(i).NameEquals Is Nothing OrElse selectVariables(i).NameEquals.AsClause Is Nothing)

                        ' Using enclosing binder for shadowing check, because range variables brought in scope by the current binder
                        ' are no longer in scope after the Select. So, it is fine to shadow them. 
                        Dim rangeVar As RangeVariableSymbol = Me.BindExpressionRangeVariable(selectVariables(i), True, Me.ContainingBinder, declaredNames, selectors(i), diagnostics)
                        Debug.Assert(rangeVar IsNot Nothing)
                        rangeVariables(i) = rangeVar
                    Next

                    AssertDeclaredNames(declaredNames, rangeVariables.AsImmutableOrNull())

                    Dim fields = New AnonymousTypeField(selectVariables.Count - 1) {}

                    For i As Integer = 0 To rangeVariables.Length - 1
                        Dim rangeVar As RangeVariableSymbol = rangeVariables(i)
                        fields(i) = New AnonymousTypeField(
                            rangeVar.Name, rangeVar.Type, rangeVar.Syntax.GetLocation(), isKeyOrByRef:=True)
                    Next

                    selector = BindAnonymousObjectCreationExpression(selectorSyntax,
                                                                     New AnonymousTypeDescriptor(fields.AsImmutableOrNull(),
                                                                                                 selectorSyntax.QueryClauseKeywordOrRangeVariableIdentifier.GetLocation(),
                                                                                                 True),
selectors.AsImmutableOrNull(),
                                                                     diagnostics).MakeCompilerGenerated()

                    declaredRangeVariables = rangeVariables.AsImmutableOrNull()
                End If

                Debug.Assert(Not declaredRangeVariables.IsDefault AndAlso selectorSyntax IsNot Nothing)
                Return selector
            End Function

            ''' <summary>
            ''' Bind ExpressionRangeVariableSyntax in context of this binder.
            ''' </summary>
            Private Function BindExpressionRangeVariable(
                item As ExpressionRangeVariableSyntax,
                requireRangeVariable As Boolean,
                shadowingCheckBinder As Binder,
                declaredNames As HashSet(Of String),
                <Out()> ByRef selector As BoundExpression,
                diagnostics As DiagnosticBag
            ) As RangeVariableSymbol
                Debug.Assert(selector Is Nothing)

                Dim variableName As VariableNameEqualsSyntax = item.NameEquals

                ' Figure out the name of the new range variable
                Dim rangeVarName As String = Nothing
                Dim rangeVarNameSyntax As SyntaxToken = Nothing
                Dim targetVariableType As TypeSymbol = Nothing

                If variableName IsNot Nothing Then
                    rangeVarNameSyntax = variableName.Identifier.Identifier
                    rangeVarName = rangeVarNameSyntax.ValueText

                    ' Deal with AsClauseOpt and various modifiers.
                    If variableName.AsClause IsNot Nothing Then
                        targetVariableType = DecodeModifiedIdentifierType(variableName.Identifier,
                                                                          variableName.AsClause,
                                                                          Nothing,
                                                                          Nothing,
                                                                          diagnostics,
                                                                          ModifiedIdentifierTypeDecoderContext.LocalType Or
                                                                          ModifiedIdentifierTypeDecoderContext.QueryRangeVariableType)

                    ElseIf variableName.Identifier.Nullable.Node IsNot Nothing Then
                        ReportDiagnostic(diagnostics, variableName.Identifier.Nullable, ERRID.ERR_NullableTypeInferenceNotSupported)
                    End If
                Else
                    ' Infer the name from expression
                    Dim failedToInferFromXmlName As XmlNameSyntax = Nothing
                    Dim nameToken As SyntaxToken = item.Expression.ExtractAnonymousTypeMemberName(failedToInferFromXmlName)

                    If nameToken.Kind <> SyntaxKind.None Then
                        rangeVarNameSyntax = nameToken
                        rangeVarName = rangeVarNameSyntax.ValueText
                    ElseIf requireRangeVariable Then
                        If failedToInferFromXmlName IsNot Nothing Then
                            ReportDiagnostic(diagnostics, failedToInferFromXmlName.LocalName, ERRID.ERR_QueryAnonTypeFieldXMLNameInference)
                        Else
                            ReportDiagnostic(diagnostics, item.Expression, ERRID.ERR_QueryAnonymousTypeFieldNameInference)
                        End If
                    End If
                End If

                If rangeVarName IsNot Nothing AndAlso rangeVarName.Length = 0 Then
                    ' Empty string must have been a syntax error. 
                    rangeVarName = Nothing
                End If

                selector = BindValue(item.Expression, diagnostics)

                If targetVariableType Is Nothing Then
                    selector = MakeRValue(selector, diagnostics)
                Else
                    selector = ApplyImplicitConversion(item.Expression, targetVariableType, selector, diagnostics)
                End If

                Dim rangeVar As RangeVariableSymbol = Nothing

                If rangeVarName IsNot Nothing Then

                    rangeVar = RangeVariableSymbol.Create(Me, rangeVarNameSyntax, selector.Type)

                    ' Note what we are doing here:
                    ' We are creating BoundRangeVariableAssignment before doing any shadowing checks
                    ' so that SemanticModel can find the declared symbol, but, if the variable will conflict with another
                    ' variable in the same child scope, we will not add it to the scope. Instead, we create special
                    ' error recovery range variable symbol and add it to the scope at the same place, making sure 
                    ' that the earlier declared range variable wins during name lookup.
                    ' As an alternative, we could still add the original range variable symbol to the scope and then,
                    ' while we are binding the rest of the query in error recovery mode, references to the name would 
                    ' cause ambiguity. However, this could negatively affect IDE experience. Also, as we build an
                    ' Anonymous Type for the compound range variables, we would end up with a type with duplicate members,
                    ' which could cause problems elsewhere.

                    selector = New BoundRangeVariableAssignment(item, rangeVar, selector, selector.Type)

                    Dim doErrorRecovery As Boolean = False

                    If declaredNames IsNot Nothing AndAlso Not declaredNames.Add(rangeVarName) Then
                        Debug.Assert(item.Parent.Kind = SyntaxKind.SelectClause OrElse item.Parent.Kind = SyntaxKind.GroupByClause)
                        ReportDiagnostic(diagnostics, rangeVarNameSyntax, ERRID.ERR_QueryDuplicateAnonTypeMemberName1, rangeVarName)
                        doErrorRecovery = True  ' Shouldn't add to the scope.
                    Else
                        shadowingCheckBinder.VerifyRangeVariableName(rangeVar, rangeVarNameSyntax, diagnostics)

                        If item.Parent.Kind = SyntaxKind.LetClause AndAlso ShadowsRangeVariableInTheChildScope(shadowingCheckBinder, rangeVar) Then
                            ' We are about to add this range variable to the current child scope.
                            ' Shadowing error was reported earlier.
                            doErrorRecovery = True  ' Shouldn't add to the scope.
                        End If
                    End If

                    If doErrorRecovery Then
                        If requireRangeVariable Then
                            rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, rangeVar.Syntax, selector.Type)
                        Else
                            rangeVar = Nothing
                        End If
                    End If

                ElseIf requireRangeVariable Then
                    Debug.Assert(rangeVar Is Nothing)
                    rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, item, selector.Type)
                End If

                Debug.Assert(selector IsNot Nothing)
                Debug.Assert(Not requireRangeVariable OrElse rangeVar IsNot Nothing)
                Return rangeVar
            End Function

            ''' <summary>
            ''' Bind Let operator selector for a particular ExpressionRangeVariableSyntax.
            ''' Takes care of "carrying over" of previously declared range variables as well as introduction of the new one.
            ''' </summary>
            Public Function BindLetClauseVariableSelector(
                variable As ExpressionRangeVariableSyntax,
                operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                <Out()> ByRef declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(declaredRangeVariables.IsDefault)

                Dim selector As BoundExpression = Nothing

                ' Using selector binder for shadowing checks because all its range variables stay in scope
                ' after the operator.
                Dim rangeVar As RangeVariableSymbol = Me.BindExpressionRangeVariable(variable, True, Me, Nothing, selector, diagnostics)
                Debug.Assert(rangeVar IsNot Nothing)

                declaredRangeVariables = ImmutableArray.Create(rangeVar)

                If _rangeVariables.Length > 0 Then
                    ' Need to build an Anonymous Type.

                    ' If it is not the last variable in the list, we simply combine source's
                    ' compound variable (an instance of its Anonymous Type) with our new variable, 
                    ' creating new compound variable of nested Anonymous Type.

                    ' In some scenarios, it is safe to leave compound variable in nested form when there is an
                    ' operator down the road that does its own projection (Select, Group By, ...). 
                    ' All following operators have to take an Anonymous Type in both cases and, since there is no way to
                    ' restrict the shape of the Anonymous Type in method's declaration, the operators should be
                    ' insensitive to the shape of the Anonymous Type.
                    selector = BuildJoinSelector(variable,
                                                 (variable Is DirectCast(operatorsEnumerator.Current, LetClauseSyntax).Variables.Last() AndAlso
                                                      MustProduceFlatCompoundVariable(operatorsEnumerator)),
                                                 diagnostics,
                                                 rangeVar, selector)

                Else
                    ' Easy case, no need to build an Anonymous Type.
                    Debug.Assert(_rangeVariables.Length = 0)
                End If

                Debug.Assert(Not declaredRangeVariables.IsDefault)
                Return selector
            End Function


            ''' <summary>
            ''' Bind body of a lambda representing first Select operator selector for an aggregate clause in context of this binder.
            ''' </summary>
            Public Function BindAggregateClauseFirstSelector(
                aggregate As AggregateClauseSyntax,
                operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                rangeVariablesPart1 As ImmutableArray(Of RangeVariableSymbol),
                rangeVariablesPart2 As ImmutableArray(Of RangeVariableSymbol),
                <Out()> ByRef declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out()> ByRef group As BoundQueryClauseBase,
                <Out()> ByRef intoBinder As IntoClauseDisallowGroupReferenceBinder,
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(operatorsEnumerator.Current Is aggregate)
                Debug.Assert(declaredRangeVariables.IsDefault)
                Debug.Assert((rangeVariablesPart1.Length = 0) = (rangeVariablesPart2 = _rangeVariables))
                Debug.Assert((rangeVariablesPart2.Length = 0) = (rangeVariablesPart1 = _rangeVariables))
                Debug.Assert(_lambdaSymbol.ParameterCount = If(rangeVariablesPart2.Length = 0, 1, 2))
                Debug.Assert(_rangeVariables.Length = rangeVariablesPart1.Length + rangeVariablesPart2.Length)

                ' Let's interpret our group.
                Dim groupAdditionalOperators As SyntaxList(Of QueryClauseSyntax).Enumerator = aggregate.AdditionalQueryOperators.GetEnumerator()

                ' Note, this call can advance [groupAdditionalOperators] enumerator if it absorbs the following Let or Select.
                group = BindCollectionRangeVariables(aggregate, Nothing, aggregate.Variables, groupAdditionalOperators, diagnostics)
                group = BindSubsequentQueryOperators(group, groupAdditionalOperators, diagnostics)

                ' Now, deal with aggregate functions.
                Dim aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax) = aggregate.AggregationVariables
                Dim letSelector As BoundExpression
                Dim groupRangeVar As RangeVariableSymbol = Nothing

                Select Case aggregationVariables.Count
                    Case 0
                        ' Malformed syntax tree.
                        Debug.Assert(aggregationVariables.Count > 0, "Malformed syntax tree.")
                        declaredRangeVariables = ImmutableArray.Create(Of RangeVariableSymbol)(
                                       RangeVariableSymbol.CreateForErrorRecovery(Me,
                                                                                  aggregate,
                                                                                  ErrorTypeSymbol.UnknownResultType))

                        letSelector = BadExpression(aggregate, group, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()

                        intoBinder = New IntoClauseDisallowGroupReferenceBinder(New QueryLambdaBinder(_lambdaSymbol,
                                                                                                      ImmutableArray(Of RangeVariableSymbol).Empty),
                                                                                group, group.RangeVariables, group.CompoundVariableType,
_rangeVariables.Concat(group.RangeVariables))

                    Case 1
                        ' Simple case - one aggregate function.
                        '
                        ' FROM a in AA              FROM a in AA
                        ' AGGREGATE b in a.BB  =>   LET count = (FROM b IN a.BB).Count()
                        ' INTO Count()
                        '
                        intoBinder = New IntoClauseDisallowGroupReferenceBinder(New QueryLambdaBinder(_lambdaSymbol,
                                                                                                      ImmutableArray(Of RangeVariableSymbol).Empty),
                                                                                group, group.RangeVariables, group.CompoundVariableType,
_rangeVariables.Concat(group.RangeVariables))

                        Dim compoundKeyReferencePart1 As BoundExpression
                        Dim keysRangeVariablesPart1 As ImmutableArray(Of RangeVariableSymbol)
                        Dim compoundKeyReferencePart2 As BoundExpression
                        Dim keysRangeVariablesPart2 As ImmutableArray(Of RangeVariableSymbol)
                        Dim letSelectorParam As BoundLambdaParameterSymbol

                        If rangeVariablesPart1.Length > 0 Then
                            letSelectorParam = DirectCast(_lambdaSymbol.Parameters(0), BoundLambdaParameterSymbol)
                            compoundKeyReferencePart1 = New BoundParameter(letSelectorParam.Syntax, letSelectorParam, False, letSelectorParam.Type).MakeCompilerGenerated()
                            keysRangeVariablesPart1 = rangeVariablesPart1

                            If rangeVariablesPart2.Length > 0 Then
                                letSelectorParam = DirectCast(_lambdaSymbol.Parameters(1), BoundLambdaParameterSymbol)
                                compoundKeyReferencePart2 = New BoundParameter(letSelectorParam.Syntax, letSelectorParam, False, letSelectorParam.Type).MakeCompilerGenerated()
                                keysRangeVariablesPart2 = rangeVariablesPart2
                            Else
                                compoundKeyReferencePart2 = Nothing
                                keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                            End If
                        ElseIf rangeVariablesPart2.Length > 0 Then
                            letSelectorParam = DirectCast(_lambdaSymbol.Parameters(1), BoundLambdaParameterSymbol)
                            compoundKeyReferencePart1 = New BoundParameter(letSelectorParam.Syntax, letSelectorParam, False, letSelectorParam.Type).MakeCompilerGenerated()
                            keysRangeVariablesPart1 = rangeVariablesPart2
                            compoundKeyReferencePart2 = Nothing
                            keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                        Else
                            compoundKeyReferencePart1 = Nothing
                            keysRangeVariablesPart1 = ImmutableArray(Of RangeVariableSymbol).Empty
                            compoundKeyReferencePart2 = Nothing
                            keysRangeVariablesPart2 = ImmutableArray(Of RangeVariableSymbol).Empty
                        End If

                        letSelector = intoBinder.BindIntoSelector(aggregate,
                                                                  _rangeVariables,
                                                                  compoundKeyReferencePart1,
                                                                  keysRangeVariablesPart1,
                                                                  compoundKeyReferencePart2,
                                                                  keysRangeVariablesPart2,
                                                                  Nothing,
                                                                  aggregationVariables,
                                                                  MustProduceFlatCompoundVariable(operatorsEnumerator),
                                                                  declaredRangeVariables,
                                                                  diagnostics)

                    Case Else

                        ' Complex case:
                        '
                        ' FROM a in AA              FROM a in AA
                        ' AGGREGATE b in a.BB  =>   LET Group = (FROM b IN a.BB)
                        ' INTO Count(),             Select a, Count=Group.Count(), Sum=Group.Sum(b=>b)
                        '      Sum(b)
                        '

                        ' Handle selector for the Let.
                        groupRangeVar = RangeVariableSymbol.CreateCompilerGenerated(Me, aggregate, StringConstants.Group, group.Type)

                        If _rangeVariables.Length = 0 Then
                            letSelector = group
                        Else
                            letSelector = BuildJoinSelector(aggregate,
                                                            mustProduceFlatCompoundVariable:=False, diagnostics:=diagnostics,
                                                            rangeVarOpt:=groupRangeVar, rangeVarValueOpt:=group)
                        End If

                        declaredRangeVariables = ImmutableArray.Create(Of RangeVariableSymbol)(groupRangeVar)
                        intoBinder = Nothing
                End Select

                Return letSelector
            End Function

            ''' <summary>
            ''' Bind Join/From selector that absorbs following Select/Let in context of this binder.
            ''' </summary>
            Public Function BindAbsorbingJoinSelector(
                absorbNextOperator As QueryClauseSyntax,
                operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                leftRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                rightRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out> ByRef joinSelectorDeclaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out> ByRef group As BoundQueryClauseBase,
                <Out> ByRef intoBinder As IntoClauseDisallowGroupReferenceBinder,
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(joinSelectorDeclaredRangeVariables.IsDefault)
                Debug.Assert(absorbNextOperator Is operatorsEnumerator.Current)

                group = Nothing
                intoBinder = Nothing

                Dim joinSelector As BoundExpression

                Select Case absorbNextOperator.Kind
                    Case SyntaxKind.SelectClause
                        ' Absorb Select
                        Dim [select] = DirectCast(absorbNextOperator, SelectClauseSyntax)
                        joinSelector = BindSelectClauseSelector([select],
                                                                operatorsEnumerator,
                                                                joinSelectorDeclaredRangeVariables,
                                                                diagnostics)

                    Case SyntaxKind.LetClause
                        ' Absorb Let
                        Dim [let] = DirectCast(absorbNextOperator, LetClauseSyntax)
                        Debug.Assert([let].Variables.Count > 0)

                        joinSelector = BindLetClauseVariableSelector([let].Variables.First,
                                                                     operatorsEnumerator,
                                                                     joinSelectorDeclaredRangeVariables,
                                                                     diagnostics)
                    Case SyntaxKind.AggregateClause
                        ' Absorb Aggregate
                        Dim aggregate = DirectCast(absorbNextOperator, AggregateClauseSyntax)
                        joinSelector = BindAggregateClauseFirstSelector(aggregate,
                                                                        operatorsEnumerator,
                                                                        leftRangeVariables,
                                                                        rightRangeVariables,
                                                                        joinSelectorDeclaredRangeVariables,
                                                                        group,
                                                                        intoBinder,
                                                                        diagnostics)
                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(absorbNextOperator.Kind)

                End Select

                Return joinSelector
            End Function
 
            ''' <summary>
            ''' Bind Join/Let like and mixed selector in context of this binder.
            ''' 
            ''' Join like selector: Function(a, b) New With {a, b}
            ''' 
            ''' Let like selector: Function(a) New With {a, letExpressionRangeVariable}
            ''' 
            ''' Mixed selector: Function(a, b) New With {a, b, letExpressionRangeVariable}
            ''' </summary>
            Public Function BuildJoinSelector(
                syntax As VisualBasicSyntaxNode,
                mustProduceFlatCompoundVariable As Boolean,
                diagnostics As DiagnosticBag,
                Optional rangeVarOpt As RangeVariableSymbol = Nothing,
                Optional rangeVarValueOpt As BoundExpression = Nothing
            ) As BoundExpression
                Debug.Assert(_rangeVariables.Length > 0)
                Debug.Assert((rangeVarOpt Is Nothing) = (rangeVarValueOpt Is Nothing))
                Debug.Assert(rangeVarOpt IsNot Nothing OrElse _lambdaSymbol.ParameterCount > 1)

                Dim selectors As BoundExpression()
                Dim fields As AnonymousTypeField()

                ' In some scenarios, it is safe to leave compound variable in nested form when there is an
                ' operator down the road that does its own projection (Select, Group By, ...). 
                ' All following operators have to take an Anonymous Type in both cases and, since there is no way to
                ' restrict the shape of the Anonymous Type in method's declaration, the operators should be
                ' insensitive to the shape of the Anonymous Type.
                Dim lastIndex As Integer
                Dim sizeIncrease As Integer = If(rangeVarOpt Is Nothing, 0, 1)

                Debug.Assert(sizeIncrease + _rangeVariables.Length > 1)

                ' Note, the special case for [sourceRangeVariables.Count = 1] helps in the
                ' following scenario:
                '     From x In Xs Select x+1 From y In Ys Let z= Zz, ...
                ' Selector lambda should be:
                '     Function(unnamed, y) New With {y, .z=Zz}
                ' The lambda has two parameters, but we have only one range variable that should be carried over.
                ' If we were simply copying lambda's parameters to the Anonymous Type instance, we would 
                ' copy data that aren't needed (value of the first parameter should be dropped).
                If _rangeVariables.Length = 1 OrElse mustProduceFlatCompoundVariable Then
                    ' Need to flatten
                    lastIndex = _rangeVariables.Length + sizeIncrease - 1
                    selectors = New BoundExpression(lastIndex) {}
                    fields = New AnonymousTypeField(lastIndex) {}

                    For j As Integer = 0 To _rangeVariables.Length - 1
                        Dim leftVar As RangeVariableSymbol = _rangeVariables(j)
                        selectors(j) = New BoundRangeVariable(leftVar.Syntax, leftVar, leftVar.Type).MakeCompilerGenerated()
                        fields(j) = New AnonymousTypeField(leftVar.Name, leftVar.Type, leftVar.Syntax.GetLocation(), isKeyOrByRef:=True)
                    Next
                Else
                    ' Nesting ...
                    Debug.Assert(_rangeVariables.Length > 1)

                    Dim parameters As ImmutableArray(Of BoundLambdaParameterSymbol) = _lambdaSymbol.Parameters.As(Of BoundLambdaParameterSymbol)

                    lastIndex = parameters.Length + sizeIncrease - 1
                    selectors = New BoundExpression(lastIndex) {}
                    fields = New AnonymousTypeField(lastIndex) {}

                    For j As Integer = 0 To parameters.Length - 1
                        Dim param As BoundLambdaParameterSymbol = parameters(j)
                        selectors(j) = New BoundParameter(param.Syntax, param, False, param.Type).MakeCompilerGenerated()
                        fields(j) = New AnonymousTypeField(param.Name, param.Type, param.Syntax.GetLocation(), isKeyOrByRef:=True)
                    Next
                End If

                If rangeVarOpt IsNot Nothing Then
                    selectors(lastIndex) = rangeVarValueOpt
                    fields(lastIndex) = New AnonymousTypeField(
                        rangeVarOpt.Name, rangeVarOpt.Type, rangeVarOpt.Syntax.GetLocation(), isKeyOrByRef:=True)
                End If

                Return BindAnonymousObjectCreationExpression(syntax,
                                                             New AnonymousTypeDescriptor(fields.AsImmutableOrNull(),
                                                                                         syntax.QueryClauseKeywordOrRangeVariableIdentifier.GetLocation(),
                                                                                         True),
selectors.AsImmutableOrNull(),
                                                             diagnostics).MakeCompilerGenerated()
            End Function

            ''' <summary>
            ''' Bind key selectors for a Join/Group Join operator.
            ''' </summary>
            Public Shared Sub BindJoinKeys(
                parentBinder As Binder,
                join As JoinClauseSyntax,
                outer As BoundQueryClauseBase,
                inner As BoundQueryClauseBase,
                joinSelectorRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out()> ByRef outerKeyLambda As BoundQueryLambda,
                <Out()> ByRef outerKeyBinder As QueryLambdaBinder,
                <Out()> ByRef innerKeyLambda As BoundQueryLambda,
                <Out()> ByRef innerKeyBinder As QueryLambdaBinder,
                diagnostics As DiagnosticBag
            )
                Debug.Assert(outerKeyLambda Is Nothing)
                Debug.Assert(outerKeyBinder Is Nothing)
                Debug.Assert(innerKeyLambda Is Nothing)
                Debug.Assert(innerKeyBinder Is Nothing)
                Debug.Assert(joinSelectorRangeVariables.SequenceEqual(outer.RangeVariables.Concat(inner.RangeVariables)))

                Dim outerKeyParam As BoundLambdaParameterSymbol = parentBinder.CreateQueryLambdaParameterSymbol(
                                                                                                   GetQueryLambdaParameterName(outer.RangeVariables), 0,
                                                                                                   outer.CompoundVariableType,
                                                                                                   join, outer.RangeVariables)

                Dim outerKeyLambdaSymbol = parentBinder.CreateQueryLambdaSymbol(LambdaUtilities.GetJoinLeftLambdaBody(join),
                                                                                SynthesizedLambdaKind.JoinLeftQueryLambda,
                                                                                ImmutableArray.Create(outerKeyParam))
                outerKeyBinder = New QueryLambdaBinder(outerKeyLambdaSymbol, joinSelectorRangeVariables)

                Dim innerKeyParam As BoundLambdaParameterSymbol = parentBinder.CreateQueryLambdaParameterSymbol(
                                                                                                   GetQueryLambdaParameterName(inner.RangeVariables), 0,
                                                                                                   inner.CompoundVariableType,
                                                                                                   join, inner.RangeVariables)

                Dim innerKeyLambdaSymbol = parentBinder.CreateQueryLambdaSymbol(LambdaUtilities.GetJoinRightLambdaBody(join),
                                                                                SynthesizedLambdaKind.JoinRightQueryLambda,
                                                                                ImmutableArray.Create(innerKeyParam))

                innerKeyBinder = New QueryLambdaBinder(innerKeyLambdaSymbol, joinSelectorRangeVariables)

                Dim suppressDiagnostics As DiagnosticBag = Nothing

                Dim sideDeterminator As New JoinConditionSideDeterminationVisitor(outer.RangeVariables, inner.RangeVariables)

                Dim joinConditions As SeparatedSyntaxList(Of JoinConditionSyntax) = join.JoinConditions
                Dim outerKey As BoundExpression
                Dim innerKey As BoundExpression
                Dim keysAreGood As Boolean

                If joinConditions.Count = 0 Then
                    ' Malformed tree.
                    Debug.Assert(joinConditions.Count > 0, "Malformed syntax tree.")
                    outerKey = BadExpression(join, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                    innerKey = BadExpression(join, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                    keysAreGood = False

                ElseIf joinConditions.Count = 1 Then
                    ' Simple case, no need to build an Anonymous Type.
                    outerKey = Nothing
                    innerKey = Nothing
                    keysAreGood = BindJoinCondition(joinConditions(0),
                                                    outerKeyBinder,
                                                    outer.RangeVariables,
                                                    outerKey,
                                                    innerKeyBinder,
                                                    inner.RangeVariables,
                                                    innerKey,
                                                    sideDeterminator,
                                                    diagnostics,
                                                    suppressDiagnostics)

                    If Not keysAreGood Then
                        If outerKey.Type IsNot ErrorTypeSymbol.UnknownResultType Then
                            outerKey = BadExpression(outerKey.Syntax, outerKey, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                        End If
                        If innerKey.Type IsNot ErrorTypeSymbol.UnknownResultType Then
                            innerKey = BadExpression(innerKey.Syntax, innerKey, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                        End If
                    End If
                Else
                    ' Need to build a compound key as an instance of an Anonymous Type.
                    Dim outerKeys(joinConditions.Count - 1) As BoundExpression
                    Dim innerKeys(joinConditions.Count - 1) As BoundExpression
                    keysAreGood = True

                    For i As Integer = 0 To joinConditions.Count - 1
                        If Not BindJoinCondition(joinConditions(i),
                                                 outerKeyBinder,
                                                 outer.RangeVariables,
                                                 outerKeys(i),
                                                 innerKeyBinder,
                                                 inner.RangeVariables,
                                                 innerKeys(i),
                                                 sideDeterminator,
                                                 diagnostics,
                                                 suppressDiagnostics) Then
                            keysAreGood = False
                        End If
                    Next

                    If keysAreGood Then
                        Dim fields(joinConditions.Count - 1) As AnonymousTypeField

                        For i As Integer = 0 To joinConditions.Count - 1
                            fields(i) = New AnonymousTypeField(
                                "Key" & i.ToString(), outerKeys(i).Type, joinConditions(i).GetLocation(), isKeyOrByRef:=True)
                        Next

                        Dim descriptor As New AnonymousTypeDescriptor(fields.AsImmutableOrNull(),
                                                                      join.OnKeyword.GetLocation(),
                                                                      True)

                        outerKey = outerKeyBinder.BindAnonymousObjectCreationExpression(join, descriptor, outerKeys.AsImmutableOrNull(),
                                                                                        diagnostics)
                        innerKey = innerKeyBinder.BindAnonymousObjectCreationExpression(join, descriptor, innerKeys.AsImmutableOrNull(),
                                                                                        diagnostics)
                    Else
                        outerKey = BadExpression(join, outerKeys.AsImmutableOrNull(), ErrorTypeSymbol.UnknownResultType)
                        innerKey = BadExpression(join, innerKeys.AsImmutableOrNull(), ErrorTypeSymbol.UnknownResultType)
                    End If

                    outerKey.MakeCompilerGenerated()
                    innerKey.MakeCompilerGenerated()
                End If

                If suppressDiagnostics IsNot Nothing Then
                    suppressDiagnostics.Free()
                End If

                outerKeyLambda = CreateBoundQueryLambda(outerKeyLambdaSymbol,
                                                        outer.RangeVariables,
                                                        outerKey,
                                                        exprIsOperandOfConditionalBranch:=False)

                outerKeyLambda.SetWasCompilerGenerated()

                innerKeyLambda = CreateBoundQueryLambda(innerKeyLambdaSymbol,
                                                        inner.RangeVariables,
                                                        innerKey,
                                                        exprIsOperandOfConditionalBranch:=False)

                innerKeyLambda.SetWasCompilerGenerated()

                Debug.Assert(outerKeyLambda IsNot Nothing)
                Debug.Assert(innerKeyLambda IsNot Nothing)
            End Sub

            Private Shared Function BindJoinCondition(
                joinCondition As JoinConditionSyntax,
                outerKeyBinder As QueryLambdaBinder,
                outerRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out()> ByRef outerKey As BoundExpression,
                innerKeyBinder As QueryLambdaBinder,
                innerRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                <Out()> ByRef innerKey As BoundExpression,
                sideDeterminator As JoinConditionSideDeterminationVisitor,
                diagnostics As DiagnosticBag,
                ByRef suppressDiagnostics As DiagnosticBag
            ) As Boolean
                Debug.Assert(outerKey Is Nothing AndAlso innerKey Is Nothing)

                ' For a JoinConditionSyntax (<left expr> Equals <right expr>), we are going to bind
                ' left side with outerKeyBinder, which has all range variables from outer and from inner
                ' in scope. Then, we try to figure out to which join side the left key actually belongs.
                ' If it turns out that it belongs to the inner side, we will bind the left again,
                ' but now using the innerKeyBinder. The reason why we need to rebind expressions is that
                ' innerKeyBinder, outerKeyBinder are each associated with different lambda symbols and
                ' the binding process might capture the lambda in the tree or in symbols referenced by
                ' the tree, so we can't safely move bound nodes between lambdas. Usually expressions are
                ' small and very simple, so rebinding them shouldn't create a noticeable performance impact.
                ' If that will not turn out to be the case, we might try to optimize by tracking if someone
                ' requested ContainingMember from outerKeyBinder while an expression is being bound and
                ' rebind only in that case. 
                ' Similarly, in some scenarios we bind right key with innerKeyBinder first and then rebind 
                ' it with outerKeyBinder.

                Dim keysAreGood As Boolean

                Dim left As BoundExpression = outerKeyBinder.BindRValue(joinCondition.Left, diagnostics)

                Dim leftSide As JoinConditionSideDeterminationVisitor.Result
                leftSide = sideDeterminator.DetermineTheSide(left, diagnostics)

                If leftSide = JoinConditionSideDeterminationVisitor.Result.Outer Then
                    outerKey = left
                    innerKey = innerKeyBinder.BindRValue(joinCondition.Right, diagnostics)
                    keysAreGood = innerKeyBinder.VerifyJoinKeys(outerKey, outerRangeVariables, leftSide,
                                                                innerKey, innerRangeVariables, sideDeterminator.DetermineTheSide(innerKey, diagnostics),
                                                                diagnostics)
                ElseIf leftSide = JoinConditionSideDeterminationVisitor.Result.Inner Then
                    ' Rebind with the inner binder.
                    If suppressDiagnostics Is Nothing Then
                        suppressDiagnostics = DiagnosticBag.GetInstance()
                    End If

                    innerKey = innerKeyBinder.BindRValue(joinCondition.Left, suppressDiagnostics)
                    Debug.Assert(leftSide = sideDeterminator.DetermineTheSide(innerKey, diagnostics))
                    outerKey = outerKeyBinder.BindRValue(joinCondition.Right, diagnostics)
                    keysAreGood = innerKeyBinder.VerifyJoinKeys(outerKey, outerRangeVariables, sideDeterminator.DetermineTheSide(outerKey, diagnostics),
                                                                innerKey, innerRangeVariables, leftSide,
                                                                diagnostics)
                Else
                    Dim right As BoundExpression = innerKeyBinder.BindRValue(joinCondition.Right, diagnostics)

                    Dim rightSide As JoinConditionSideDeterminationVisitor.Result
                    rightSide = sideDeterminator.DetermineTheSide(right, diagnostics)

                    If rightSide = JoinConditionSideDeterminationVisitor.Result.Outer Then
                        ' Rebind with the outer binder.
                        If suppressDiagnostics Is Nothing Then
                            suppressDiagnostics = DiagnosticBag.GetInstance()
                        End If

                        outerKey = outerKeyBinder.BindRValue(joinCondition.Right, suppressDiagnostics)
                        innerKey = innerKeyBinder.BindRValue(joinCondition.Left, suppressDiagnostics)
                        keysAreGood = innerKeyBinder.VerifyJoinKeys(outerKey, outerRangeVariables, rightSide,
                                                                    innerKey, innerRangeVariables, leftSide,
                                                                    diagnostics)
                    Else
                        outerKey = left
                        innerKey = right

                        keysAreGood = innerKeyBinder.VerifyJoinKeys(outerKey, outerRangeVariables, leftSide,
                                                                    innerKey, innerRangeVariables, rightSide,
                                                                    diagnostics)
                    End If

                    Debug.Assert(Not keysAreGood)
                End If

                If keysAreGood Then
                    keysAreGood = Not (outerKey.Type.IsErrorType() OrElse innerKey.Type.IsErrorType())
                End If

                If keysAreGood AndAlso Not outerKey.Type.IsSameTypeIgnoringAll(innerKey.Type) Then
                    ' Apply conversion if available.
                    Dim targetType As TypeSymbol = Nothing
                    Dim intrinsicOperatorType As SpecialType = SpecialType.None
                    Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
                    Dim operatorKind As BinaryOperatorKind = OverloadResolution.ResolveBinaryOperator(BinaryOperatorKind.Equals,
                                                                                                      outerKey, innerKey,
                                                                                                      innerKeyBinder,
                                                                                                      considerUserDefinedOrLateBound:=False,
                                                                                                      intrinsicOperatorType:=intrinsicOperatorType,
                                                                                                      userDefinedOperator:=Nothing,
                                                                                                      useSiteDiagnostics:=useSiteDiagnostics)

                    If (operatorKind And BinaryOperatorKind.Equals) <> 0 AndAlso
                       (operatorKind And Not (BinaryOperatorKind.Equals Or BinaryOperatorKind.Lifted)) = 0 AndAlso
                       intrinsicOperatorType <> SpecialType.None Then
                        ' There is an intrinsic (=) operator, use its argument type. 
                        Debug.Assert(useSiteDiagnostics.IsNullOrEmpty)
                        targetType = innerKeyBinder.GetSpecialTypeForBinaryOperator(joinCondition, outerKey.Type, innerKey.Type, intrinsicOperatorType,
                                                                                    (operatorKind And BinaryOperatorKind.Lifted) <> 0, diagnostics)
                    Else
                        ' Use dominant type.
                        Dim inferenceCollection = New TypeInferenceCollection()
                        inferenceCollection.AddType(outerKey.Type, RequiredConversion.Any, outerKey)
                        inferenceCollection.AddType(innerKey.Type, RequiredConversion.Any, innerKey)

                        Dim resultList = ArrayBuilder(Of DominantTypeData).GetInstance()
                        inferenceCollection.FindDominantType(resultList, Nothing, useSiteDiagnostics)

                        If diagnostics.Add(joinCondition, useSiteDiagnostics) Then
                            ' Suppress additional diagnostics
                            diagnostics = New DiagnosticBag()
                        End If

                        If resultList.Count = 1 Then
                            targetType = resultList(0).ResultType
                        End If

                        resultList.Free()
                    End If

                    If targetType Is Nothing Then
                        ReportDiagnostic(diagnostics, joinCondition, ERRID.ERR_EqualsTypeMismatch, outerKey.Type, innerKey.Type)
                        keysAreGood = False
                    Else
                        outerKey = outerKeyBinder.ApplyImplicitConversion(outerKey.Syntax, targetType, outerKey, diagnostics, False)
                        innerKey = innerKeyBinder.ApplyImplicitConversion(innerKey.Syntax, targetType, innerKey, diagnostics, False)
                    End If
                End If

                Debug.Assert(outerKey IsNot Nothing AndAlso innerKey IsNot Nothing)
                Return keysAreGood
            End Function


            Private Function VerifyJoinKeys(
                outerKey As BoundExpression,
                outerRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                outerSide As JoinConditionSideDeterminationVisitor.Result,
                innerKey As BoundExpression,
                innerRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                innerSide As JoinConditionSideDeterminationVisitor.Result,
                diagnostics As DiagnosticBag
            ) As Boolean
                If outerSide = JoinConditionSideDeterminationVisitor.Result.Outer AndAlso
                   innerSide = JoinConditionSideDeterminationVisitor.Result.Inner Then
                    Return True
                End If

                If outerKey.HasErrors OrElse innerKey.HasErrors Then
                    Return False
                End If

                Dim builder = PooledStringBuilder.GetInstance()

                Dim outerNames As String = BuildEqualsOperandIsBadErrorArgument(builder.Builder, outerRangeVariables)
                Dim innerNames As String = BuildEqualsOperandIsBadErrorArgument(builder.Builder, innerRangeVariables)

                builder.Free()

                If outerNames Is Nothing OrElse innerNames Is Nothing Then
                    ' Syntax errors were earlier reported.
                    Return False
                End If

                Dim errorInfo = ErrorFactory.ErrorInfo(ERRID.ERR_EqualsOperandIsBad, outerNames, innerNames)

                Select Case outerSide
                    Case JoinConditionSideDeterminationVisitor.Result.Both,
                         JoinConditionSideDeterminationVisitor.Result.Inner
                        ' Report errors on references to inner.
                        EqualsOperandIsBadErrorVisitor.Report(Me, errorInfo, innerRangeVariables, outerKey, diagnostics)

                    Case JoinConditionSideDeterminationVisitor.Result.None
                        ReportDiagnostic(diagnostics, outerKey.Syntax, errorInfo)

                    Case JoinConditionSideDeterminationVisitor.Result.Outer
                        ' This side is good.
                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(outerSide)
                End Select

                Select Case innerSide
                    Case JoinConditionSideDeterminationVisitor.Result.Both,
                         JoinConditionSideDeterminationVisitor.Result.Outer
                        ' Report errors on references to outer.
                        EqualsOperandIsBadErrorVisitor.Report(Me, errorInfo, outerRangeVariables, innerKey, diagnostics)

                    Case JoinConditionSideDeterminationVisitor.Result.None
                        ReportDiagnostic(diagnostics, innerKey.Syntax, errorInfo)

                    Case JoinConditionSideDeterminationVisitor.Result.Inner
                        ' This side is good.
                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(innerSide)
                End Select

                Return False
            End Function

            Private Shared Function BuildEqualsOperandIsBadErrorArgument(
                builder As System.Text.StringBuilder,
                rangeVariables As ImmutableArray(Of RangeVariableSymbol)
            ) As String
                builder.Clear()

                Dim i As Integer

                For i = 0 To rangeVariables.Length - 1
                    If Not rangeVariables(i).Name.StartsWith("$"c, StringComparison.Ordinal) Then
                        builder.Append("'"c)
                        builder.Append(rangeVariables(i).Name)
                        builder.Append("'"c)
                        i += 1
                        Exit For
                    End If
                Next

                For i = i To rangeVariables.Length - 1
                    If Not rangeVariables(i).Name.StartsWith("$"c, StringComparison.Ordinal) Then
                        builder.Append(","c)
                        builder.Append(" "c)
                        builder.Append("'"c)
                        builder.Append(rangeVariables(i).Name)
                        builder.Append("'"c)
                    End If
                Next

                If builder.Length = 0 Then
                    Return Nothing
                End If

                Return builder.ToString()
            End Function

            ''' <summary>
            ''' Helper visitor to determine what join sides are referenced by an expression.
            ''' </summary>
            Private Class JoinConditionSideDeterminationVisitor
                Inherits BoundTreeWalkerWithStackGuardWithoutRecursionOnTheLeftOfBinaryOperator

                <Flags()>
                Public Enum Result
                    None = 0
                    Outer = 1
                    Inner = 2
                    Both = Outer Or Inner
                End Enum

                Private ReadOnly _outerRangeVariables As ImmutableArray(Of Object) 'ImmutableArray(Of RangeVariableSymbol)
                Private ReadOnly _innerRangeVariables As ImmutableArray(Of Object) 'ImmutableArray(Of RangeVariableSymbol)
                Private _side As Result

                Public Sub New(
                    outerRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                    innerRangeVariables As ImmutableArray(Of RangeVariableSymbol)
                )
                    _outerRangeVariables = StaticCast(Of Object).From(outerRangeVariables)
                    _innerRangeVariables = StaticCast(Of Object).From(innerRangeVariables)
                End Sub

                Public Function DetermineTheSide(node As BoundExpression, diagnostics As DiagnosticBag) As Result
                    _side = Result.None
                    Try
                        Visit(node)
                    Catch ex As CancelledByStackGuardException
                        ex.AddAnError(diagnostics)
                    End Try

                    Return _side
                End Function

                Public Overrides Function VisitRangeVariable(node As BoundRangeVariable) As BoundNode
                    Dim rangeVariable As RangeVariableSymbol = node.RangeVariable

                    If _outerRangeVariables.IndexOf(rangeVariable, ReferenceEqualityComparer.Instance) >= 0 Then
                        _side = _side Or Result.Outer
                    ElseIf _innerRangeVariables.IndexOf(rangeVariable, ReferenceEqualityComparer.Instance) >= 0 Then
                        _side = _side Or Result.Inner
                    End If

                    Return node
                End Function
            End Class


            ''' <summary>
            ''' Helper visitor to report query specific errors for an operand of an Equals expression.
            ''' </summary>
            Private Class EqualsOperandIsBadErrorVisitor
                Inherits BoundTreeWalkerWithStackGuardWithoutRecursionOnTheLeftOfBinaryOperator

                Private ReadOnly _binder As Binder
                Private ReadOnly _errorInfo As DiagnosticInfo
                Private ReadOnly _diagnostics As DiagnosticBag
                Private ReadOnly _badRangeVariables As ImmutableArray(Of Object) 'ImmutableArray(Of RangeVariableSymbol)

                Private Sub New(
                    binder As Binder,
                    errorInfo As DiagnosticInfo,
                    badRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                    diagnostics As DiagnosticBag
                )
                    _badRangeVariables = StaticCast(Of Object).From(badRangeVariables)
                    _binder = binder
                    _diagnostics = diagnostics
                    _errorInfo = errorInfo
                End Sub

                Public Shared Sub Report(
                    binder As Binder,
                    errorInfo As DiagnosticInfo,
                    badRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                    node As BoundExpression,
                    diagnostics As DiagnosticBag
                )
                    Dim v As New EqualsOperandIsBadErrorVisitor(binder, errorInfo, badRangeVariables, diagnostics)
                    Try
                        v.Visit(node)
                    Catch ex As CancelledByStackGuardException
                        ex.AddAnError(diagnostics)
                    End Try
                End Sub

                Public Overrides Function VisitRangeVariable(node As BoundRangeVariable) As BoundNode
                    Dim rangeVariable As RangeVariableSymbol = node.RangeVariable

                    If _badRangeVariables.IndexOf(rangeVariable, ReferenceEqualityComparer.Instance) >= 0 Then
                        ReportDiagnostic(_diagnostics, node.Syntax, _errorInfo)
                    End If

                    Return node
                End Function

            End Class
        End Class

    End Class

End Namespace

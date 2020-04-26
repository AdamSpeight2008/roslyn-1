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
        ''' Knows how to bind FunctionAggregationSyntax and GroupAggregationSyntax
        ''' within particular [Into] clause. 
        ''' 
        ''' Also implements Lookup/LookupNames methods to make sure that lookup without 
        ''' container type, uses type of the group as the container type.
        ''' </summary>
        Private Class IntoClauseBinder
            Inherits Binder

            Protected ReadOnly m_GroupReference As BoundExpression
            Private ReadOnly _groupRangeVariables As ImmutableArray(Of RangeVariableSymbol)
            Private ReadOnly _groupCompoundVariableType As TypeSymbol
            Private ReadOnly _aggregationArgumentRangeVariables As ImmutableArray(Of RangeVariableSymbol)

            Public Sub New(
                parent As Binder,
                groupReference As BoundExpression,
                groupRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                groupCompoundVariableType As TypeSymbol,
                aggregationArgumentRangeVariables As ImmutableArray(Of RangeVariableSymbol)
            )
                MyBase.New(parent)
                m_GroupReference = groupReference
                _groupRangeVariables = groupRangeVariables
                _groupCompoundVariableType = groupCompoundVariableType
                _aggregationArgumentRangeVariables = aggregationArgumentRangeVariables
            End Sub

            Friend Overrides Function BindGroupAggregationExpression(
                group As GroupAggregationSyntax,
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Return New BoundGroupAggregation(group, m_GroupReference, m_GroupReference.Type)
            End Function

            ''' <summary>
            ''' Given aggregationVariables, bind Into selector in context of this binder.
            ''' </summary>
            Public Function BindIntoSelector(
                syntaxNode As QueryClauseSyntax,
                keysRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                compoundKeyReferencePart1 As BoundExpression,
                keysRangeVariablesPart1 As ImmutableArray(Of RangeVariableSymbol),
                compoundKeyReferencePart2 As BoundExpression,
                keysRangeVariablesPart2 As ImmutableArray(Of RangeVariableSymbol),
                declaredNames As HashSet(Of String),
                aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax),
                mustProduceFlatCompoundVariable As Boolean,
                <Out()> ByRef declaredRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                diagnostics As DiagnosticBag
            ) As BoundExpression
                Debug.Assert(declaredRangeVariables.IsDefault)
                Debug.Assert(compoundKeyReferencePart2 Is Nothing OrElse compoundKeyReferencePart1 IsNot Nothing)
                Debug.Assert((compoundKeyReferencePart1 Is Nothing) = (keysRangeVariablesPart1.Length = 0))
                Debug.Assert((compoundKeyReferencePart2 Is Nothing) = (keysRangeVariablesPart2.Length = 0))
                Debug.Assert((compoundKeyReferencePart2 Is Nothing) = (keysRangeVariables = keysRangeVariablesPart1))
                Debug.Assert(compoundKeyReferencePart1 Is Nothing OrElse keysRangeVariables.Length = keysRangeVariablesPart1.Length + keysRangeVariablesPart2.Length)

                Dim keys As Integer = keysRangeVariables.Length
                Dim selectors As BoundExpression()
                Dim fields As AnonymousTypeField()

                If declaredNames Is Nothing Then
                    declaredNames = CreateSetOfDeclaredNames(keysRangeVariables)
                Else
                    AssertDeclaredNames(declaredNames, keysRangeVariables)
                End If

                Dim intoSelector As BoundExpression
                Dim rangeVariables() As RangeVariableSymbol

                Dim fieldsToReserveForAggregationVariables As Integer = Math.Max(aggregationVariables.Count, 1)

                If keys + fieldsToReserveForAggregationVariables > 1 Then
                    ' Need to build an Anonymous Type
                    ' Add keys first.
                    If keys > 1 AndAlso Not mustProduceFlatCompoundVariable Then
                        If compoundKeyReferencePart2 Is Nothing Then
                            selectors = New BoundExpression(fieldsToReserveForAggregationVariables + 1 - 1) {}
                            fields = New AnonymousTypeField(selectors.Length - 1) {}

                            ' Using syntax of the first range variable in the source, this shouldn't create any problems.
                            fields(0) = New AnonymousTypeField(GetQueryLambdaParameterName(keysRangeVariablesPart1),
                                                               compoundKeyReferencePart1.Type,
                                                               keysRangeVariables(0).Syntax.GetLocation(), isKeyOrByRef:=True)
                            selectors(0) = compoundKeyReferencePart1
                            keys = 1
                        Else
                            selectors = New BoundExpression(fieldsToReserveForAggregationVariables + 2 - 1) {}
                            fields = New AnonymousTypeField(selectors.Length - 1) {}

                            ' Using syntax of the first range variable in the source, this shouldn't create any problems.
                            fields(0) = New AnonymousTypeField(GetQueryLambdaParameterNameLeft(keysRangeVariablesPart1),
                                                               compoundKeyReferencePart1.Type,
                                                               keysRangeVariablesPart1(0).Syntax.GetLocation(),
                                                               isKeyOrByRef:=True)
                            selectors(0) = compoundKeyReferencePart1

                            fields(1) = New AnonymousTypeField(GetQueryLambdaParameterNameRight(keysRangeVariablesPart2),
                                                               compoundKeyReferencePart2.Type,
                                                               keysRangeVariablesPart2(0).Syntax.GetLocation(),
                                                               isKeyOrByRef:=True)
                            selectors(1) = compoundKeyReferencePart2

                            keys = 2
                        End If
                    Else
                        selectors = New BoundExpression(keys + fieldsToReserveForAggregationVariables - 1) {}
                        fields = New AnonymousTypeField(selectors.Length - 1) {}

                        For i As Integer = 0 To keys - 1
                            Dim rangeVar As RangeVariableSymbol = keysRangeVariables(i)
                            fields(i) = New AnonymousTypeField(rangeVar.Name, rangeVar.Type, rangeVar.Syntax.GetLocation(), isKeyOrByRef:=True)
                            selectors(i) = New BoundRangeVariable(rangeVar.Syntax, rangeVar, rangeVar.Type).MakeCompilerGenerated()
                        Next
                    End If

                    ' Now add aggregation variables.
                    rangeVariables = New RangeVariableSymbol(fieldsToReserveForAggregationVariables - 1) {}

                    If aggregationVariables.Count = 0 Then
                        ' Malformed syntax tree.
                        Debug.Assert(aggregationVariables.Count > 0, "Malformed syntax tree.")
                        Dim rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, syntaxNode, ErrorTypeSymbol.UnknownResultType)

                        rangeVariables(0) = rangeVar
                        selectors(keys) = BadExpression(syntaxNode, m_GroupReference, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                        fields(keys) = New AnonymousTypeField(rangeVar.Name, rangeVar.Type, rangeVar.Syntax.GetLocation(), isKeyOrByRef:=True)
                    Else
                        For i As Integer = 0 To aggregationVariables.Count - 1
                            Dim selector As BoundExpression = Nothing
                            Dim rangeVar As RangeVariableSymbol = BindAggregationRangeVariable(aggregationVariables(i),
                                                                                               declaredNames, selector,
                                                                                               diagnostics)

                            Debug.Assert(rangeVar IsNot Nothing)
                            rangeVariables(i) = rangeVar

                            selectors(keys + i) = selector
                            fields(keys + i) = New AnonymousTypeField(
                                rangeVar.Name, rangeVar.Type, rangeVar.Syntax.GetLocation(), isKeyOrByRef:=True)
                        Next
                    End If

                    Debug.Assert(selectors.Length > 1)
                    intoSelector = BindAnonymousObjectCreationExpression(syntaxNode,
                                                                     New AnonymousTypeDescriptor(fields.AsImmutableOrNull(),
                                                                                                 syntaxNode.QueryClauseKeywordOrRangeVariableIdentifier.GetLocation(),
                                                                                                 True),
                                                                     selectors.AsImmutableOrNull(),
                                                                     diagnostics).MakeCompilerGenerated()
                Else
                    Debug.Assert(keys = 0)

                    Dim rangeVar As RangeVariableSymbol

                    If aggregationVariables.Count = 0 Then
                        ' Malformed syntax tree.
                        Debug.Assert(aggregationVariables.Count > 0, "Malformed syntax tree.")
                        rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, syntaxNode, ErrorTypeSymbol.UnknownResultType)
                        intoSelector = BadExpression(syntaxNode, m_GroupReference, ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                    Else
                        Debug.Assert(aggregationVariables.Count = 1)
                        intoSelector = Nothing
                        rangeVar = BindAggregationRangeVariable(aggregationVariables(0),
                                                                declaredNames, intoSelector,
                                                                diagnostics)
                        Debug.Assert(rangeVar IsNot Nothing)
                    End If

                    rangeVariables = {rangeVar}
                End If


                AssertDeclaredNames(declaredNames, rangeVariables.AsImmutableOrNull())

                declaredRangeVariables = rangeVariables.AsImmutableOrNull()
                Return intoSelector
            End Function

            Friend Overrides Function BindFunctionAggregationExpression(
                functionAggregationSyntax As FunctionAggregationSyntax,
                diagnostics As DiagnosticBag
            ) As BoundExpression
                If functionAggregationSyntax.FunctionName.GetTypeCharacter() <> TypeCharacter.None Then
                    ReportDiagnostic(diagnostics, functionAggregationSyntax.FunctionName, ERRID.ERR_TypeCharOnAggregation)
                End If

                Dim aggregationParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(_groupRangeVariables), 0,
                                                                                                      _groupCompoundVariableType,
                                                                                                      If(functionAggregationSyntax.Argument, functionAggregationSyntax),
                                                                                                      _groupRangeVariables)

                ' Note [Binder:=Me.ContainingBinder] below. We are excluding this binder from the chain of
                ' binders for an argument of the function. Do not want it to interfere in any way given
                ' its special binding and Lookup/LookupNames behavior.

                Dim aggregationLambdaSymbol = Me.ContainingBinder.CreateQueryLambdaSymbol(
                    If(LambdaUtilities.GetAggregationLambdaBody(functionAggregationSyntax), functionAggregationSyntax),
                    SynthesizedLambdaKind.AggregationQueryLambda,
                    ImmutableArray.Create(aggregationParam))

                ' Create binder for the aggregation.
                Dim aggregationBinder As New QueryLambdaBinder(aggregationLambdaSymbol, _aggregationArgumentRangeVariables)

                Dim arguments As ImmutableArray(Of BoundExpression)
                Dim aggregationLambda As BoundQueryLambda = Nothing

                If functionAggregationSyntax.Argument Is Nothing Then
                    arguments = ImmutableArray(Of BoundExpression).Empty
                Else
                    ' Bind argument as a value, conversion during overload resolution should take care of the rest (making it an RValue, etc.). 
                    Dim aggregationSelector = aggregationBinder.BindValue(functionAggregationSyntax.Argument, diagnostics)

                    aggregationLambda = CreateBoundQueryLambda(aggregationLambdaSymbol,
                                                               _groupRangeVariables,
                                                               aggregationSelector,
                                                               exprIsOperandOfConditionalBranch:=False)

                    ' Note, we are not setting ReturnType for aggregationLambdaSymbol to allow
                    ' additional conversions. This type doesn't affect type of any range variable
                    ' in the query.
                    aggregationLambda.SetWasCompilerGenerated()

                    arguments = ImmutableArray.Create(Of BoundExpression)(aggregationLambda)
                End If

                ' Now bind the aggregation call.
                Dim boundCallOrBadExpression As BoundExpression

                If m_GroupReference.Type.IsErrorType() OrElse String.IsNullOrEmpty(functionAggregationSyntax.FunctionName.ValueText) Then
                    boundCallOrBadExpression = BadExpression(functionAggregationSyntax,
                                                             ImmutableArray.Create(m_GroupReference).AddRange(arguments),
                                                             ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
                Else
                    Dim callDiagnostics As DiagnosticBag = diagnostics

                    If aggregationLambda IsNot Nothing AndAlso ShouldSuppressDiagnostics(aggregationLambda) Then
                        ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                        callDiagnostics = DiagnosticBag.GetInstance()
                    End If

                    boundCallOrBadExpression = BindQueryOperatorCall(functionAggregationSyntax, m_GroupReference,
                                                                   functionAggregationSyntax.FunctionName.ValueText,
                                                                   arguments,
                                                                   functionAggregationSyntax.FunctionName.Span,
                                                                   callDiagnostics)

                    If callDiagnostics IsNot diagnostics Then
                        callDiagnostics.Free()
                    End If
                End If

                Return New BoundQueryClause(functionAggregationSyntax, boundCallOrBadExpression,
                                            ImmutableArray(Of RangeVariableSymbol).Empty,
                                            boundCallOrBadExpression.Type,
                                            ImmutableArray.Create(Of Binder)(aggregationBinder),
                                            boundCallOrBadExpression.Type)
            End Function


            Public Overrides Sub AddLookupSymbolsInfo(nameSet As LookupSymbolsInfo, options As LookupOptions)
                If (options And (LookupOptionExtensions.ConsiderationMask Or LookupOptions.MustNotBeInstance)) <> 0 Then
                    Return
                End If

                ' Should look for group's methods only. 
                AddMemberLookupSymbolsInfo(nameSet,
                                  m_GroupReference.Type,
                                  options Or CType(LookupOptions.MethodsOnly Or LookupOptions.MustBeInstance, LookupOptions))
            End Sub

            Public Overrides Sub Lookup(lookupResult As LookupResult, name As String, arity As Integer, options As LookupOptions, <[In], Out> ByRef useSiteDiagnostics As HashSet(Of DiagnosticInfo))
                If (options And (LookupOptionExtensions.ConsiderationMask Or LookupOptions.MustNotBeInstance)) <> 0 Then
                    Return
                End If

                ' Should look for group's methods only. 
                LookupMember(lookupResult,
                             m_GroupReference.Type,
                             name,
                             arity,
                             options Or CType(LookupOptions.MethodsOnly Or LookupOptions.MustBeInstance, LookupOptions),
                             useSiteDiagnostics)
            End Sub

            ''' <summary>
            ''' Bind AggregationRangeVariableSyntax in context of this binder.
            ''' </summary>
            Public Function BindAggregationRangeVariable(
                item As AggregationRangeVariableSyntax,
                declaredNames As HashSet(Of String),
                <Out()> ByRef selector As BoundExpression,
                diagnostics As DiagnosticBag
            ) As RangeVariableSymbol
                Debug.Assert(selector Is Nothing)

                Dim variableName As VariableNameEqualsSyntax = item.NameEquals

                ' Figure out the name of the new range variable
                Dim rangeVarName As String = Nothing
                Dim rangeVarNameSyntax As SyntaxToken = Nothing

                If variableName IsNot Nothing Then
                    rangeVarNameSyntax = variableName.Identifier.Identifier
                    rangeVarName = rangeVarNameSyntax.ValueText
                    Debug.Assert(variableName.AsClause Is Nothing)

                Else
                    ' Infer the name from expression
                    Select Case item.Aggregation.Kind
                        Case SyntaxKind.GroupAggregation
                            ' AggregateClause doesn't support GroupAggregation.
                            If item.Parent.Kind <> SyntaxKind.AggregateClause Then
                                rangeVarNameSyntax = DirectCast(item.Aggregation, GroupAggregationSyntax).GroupKeyword
                                rangeVarName = rangeVarNameSyntax.ValueText
                            End If
                        Case SyntaxKind.FunctionAggregation
                            rangeVarNameSyntax = DirectCast(item.Aggregation, FunctionAggregationSyntax).FunctionName
                            rangeVarName = rangeVarNameSyntax.ValueText
                        Case Else
                            Throw ExceptionUtilities.UnexpectedValue(item.Aggregation.Kind)
                    End Select
                End If

                ' Bind the value.
                selector = BindRValue(item.Aggregation, diagnostics)

                If rangeVarName IsNot Nothing AndAlso rangeVarName.Length = 0 Then
                    ' Empty string must have been a syntax error. 
                    rangeVarName = Nothing
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
                        ReportDiagnostic(diagnostics, rangeVarNameSyntax, ERRID.ERR_QueryDuplicateAnonTypeMemberName1, rangeVarName)
                        doErrorRecovery = True  ' Shouldn't add to the scope.
                    Else
                        Me.VerifyRangeVariableName(rangeVar, rangeVarNameSyntax, diagnostics)
                    End If

                    If doErrorRecovery Then
                        rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, rangeVar.Syntax, selector.Type)
                    End If

                Else
                    Debug.Assert(rangeVar Is Nothing)
                    rangeVar = RangeVariableSymbol.CreateForErrorRecovery(Me, item, selector.Type)
                End If

                Debug.Assert(selector IsNot Nothing)

                Return rangeVar
            End Function

        End Class

        ''' <summary>
        ''' Same as IntoClauseBinder, but disallows references to GroupAggregationSyntax.
        ''' </summary>
        Private Class IntoClauseDisallowGroupReferenceBinder
            Inherits IntoClauseBinder

            Public Sub New(
                parent As Binder,
                groupReference As BoundExpression,
                groupRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                groupCompoundVariableType As TypeSymbol,
                aggregationArgumentRangeVariables As ImmutableArray(Of RangeVariableSymbol)
            )
                MyBase.New(parent, groupReference, groupRangeVariables, groupCompoundVariableType, aggregationArgumentRangeVariables)
            End Sub

            Friend Overrides Function BindGroupAggregationExpression(group As GroupAggregationSyntax, diagnostics As DiagnosticBag) As BoundExpression
                ' Parser should have reported an error.
                Return BadExpression(group, m_GroupReference, ErrorTypeSymbol.UnknownResultType)
            End Function
        End Class

    End Class

End Namespace

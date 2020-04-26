' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Function CreateQueryLambdaSymbol(
                                                  queryLambda As (Kind As SynthesizedLambdaKind, Syntax As VisualBasicSyntaxNode),
                                                  parameters As ImmutableArray(Of BoundLambdaParameterSymbol)
                                                ) As SynthesizedLambdaSymbol

            Debug.Assert(queryLambda.kind.IsQueryLambda)

            Return New SynthesizedLambdaSymbol(queryLambda.Kind,
                                               queryLambda.Syntax,
                                               parameters,
                                               LambdaSymbol.ReturnTypePendingDelegate,
                                               Me)
        End Function

        Private Shared Function CreateBoundQueryLambda(
                                                        queryLambdaSymbol As SynthesizedLambdaSymbol,
                                                        rangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                        expression As BoundExpression,
                                                        exprIsOperandOfConditionalBranch As Boolean
                                                      ) As BoundQueryLambda
            Return New BoundQueryLambda(queryLambdaSymbol.Syntax, queryLambdaSymbol, rangeVariables, expression, exprIsOperandOfConditionalBranch)
        End Function

        Friend Overridable Function BindFunctionAggregationExpression([function] As FunctionAggregationSyntax, diagnostics As DiagnosticBag) As BoundExpression
            ' Only special query binders that have enough context can bind FunctionAggregationSyntax.
            ' TODO: Do we need to report any diagnostic?
            Debug.Assert(False, "Binding out of context is unsupported!")
            Return BadExpression([function], ErrorTypeSymbol.UnknownResultType)
        End Function

        ''' <summary>
        ''' Bind a Query Expression.
        ''' This is the entry point.
        ''' </summary>
        Private Function BindQueryExpression(
                                              query As QueryExpressionSyntax,
                                              diagnostics As DiagnosticBag
                                            ) As BoundExpression

            If query.Clauses.Count < 1 Then
                ' Syntax error must have been reported
                Return BadExpression(query, ErrorTypeSymbol.UnknownResultType)
            End If

            Dim operators As SyntaxList(Of QueryClauseSyntax).Enumerator = query.Clauses.GetEnumerator()
            Dim moveResult = operators.MoveNext()
            Debug.Assert(moveResult)

            Dim current As QueryClauseSyntax = operators.Current

            Select Case current.Kind
                Case SyntaxKind.FromClause
                    Return BindFromQueryExpression(query, operators, diagnostics)

                Case SyntaxKind.AggregateClause
                    Return BindAggregateQueryExpression(query, operators, diagnostics)

                Case Else
                    ' Syntax error must have been reported
                    Return BadExpression(query, ErrorTypeSymbol.UnknownResultType)
            End Select
        End Function

        ''' <summary>
        ''' Given a result of binding of initial set of collection range variables, the source,
        ''' bind the rest of the operators in the enumerator.
        ''' 
        ''' There is a special method to bind an operator of each kind, the common thing among them is that
        ''' all of them take the result we have so far, the source, and return result of an application 
        ''' of one or two following operators. 
        ''' Some of the methods also take operators enumerator in order to be able to do a necessary look-ahead
        ''' and in some cases even to advance the enumerator themselves.
        ''' Join and From operators absorb following Select or Let, that is when the process of binding of 
        ''' a single operator actually handles two and advances the enumerator. 
        ''' </summary>
        Private Function BindSubsequentQueryOperators(
                                                       source As BoundQueryClauseBase,
                                                       operators As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                       diagnostics As DiagnosticBag
                                                     ) As BoundQueryClauseBase
            Debug.Assert(source IsNot Nothing)

            While operators.MoveNext()
                Dim current As QueryClauseSyntax = operators.Current

                Select Case current.Kind
                    Case SyntaxKind.FromClause
                        ' Note, this call can advance [operators] enumerator if it absorbs the following Let or Select.
                        source = BindFromClause(source, DirectCast(current, FromClauseSyntax), operators, diagnostics)

                    Case SyntaxKind.SelectClause
                        source = BindSelectClause(source, DirectCast(current, SelectClauseSyntax), operators, diagnostics)

                    Case SyntaxKind.LetClause
                        source = BindLetClause(source, DirectCast(current, LetClauseSyntax), operators, diagnostics)

                    Case SyntaxKind.WhereClause
                        source = BindWhereClause(source, DirectCast(current, WhereClauseSyntax), diagnostics)

                    Case SyntaxKind.SkipWhileClause
                        source = BindSkipWhileClause(source, DirectCast(current, PartitionWhileClauseSyntax), diagnostics)

                    Case SyntaxKind.TakeWhileClause
                        source = BindTakeWhileClause(source, DirectCast(current, PartitionWhileClauseSyntax), diagnostics)

                    Case SyntaxKind.DistinctClause
                        source = BindDistinctClause(source, DirectCast(current, DistinctClauseSyntax), diagnostics)

                    Case SyntaxKind.SkipClause
                        source = BindSkipClause(source, DirectCast(current, PartitionClauseSyntax), diagnostics)

                    Case SyntaxKind.TakeClause
                        source = BindTakeClause(source, DirectCast(current, PartitionClauseSyntax), diagnostics)

                    Case SyntaxKind.OrderByClause
                        source = BindOrderByClause(source, DirectCast(current, OrderByClauseSyntax), diagnostics)

                    Case SyntaxKind.SimpleJoinClause
                        ' Note, this call can advance [operators] enumerator if it absorbs the following Let or Select.
                        source = BindInnerJoinClause(source, DirectCast(current, SimpleJoinClauseSyntax), Nothing, operators, diagnostics)

                    Case SyntaxKind.GroupJoinClause
                        source = BindGroupJoinClause(source, DirectCast(current, GroupJoinClauseSyntax), Nothing, operators, diagnostics)

                    Case SyntaxKind.GroupByClause
                        source = BindGroupByClause(source, DirectCast(current, GroupByClauseSyntax), diagnostics)

                    Case SyntaxKind.AggregateClause
                        source = BindAggregateClause(source, DirectCast(current, AggregateClauseSyntax), operators, diagnostics)

                    Case SyntaxKind.ZipClause
                        ' Note, this call can advance [operators] enumerator if it absorbs the following Let or Select.
                        source = BindZipClause(source, DirectCast(current, ZipClauseSyntax), operators, diagnostics)

                    Case Else
                        Throw ExceptionUtilities.UnexpectedValue(current.Kind)
                End Select
            End While

            Return source
        End Function

        Private Shared Function ShouldSuppressDiagnostics(lambda As BoundQueryLambda) As Boolean
            If lambda.HasErrors Then Return True

            For Each param As ParameterSymbol In lambda.LambdaSymbol.Parameters
                If param.Type.IsErrorType() Then Return True
            Next

            Dim bodyType As TypeSymbol = lambda.Expression.Type
            Return bodyType IsNot Nothing AndAlso bodyType.IsErrorType()
        End Function

        ''' <summary>
        ''' Apply "conversion" to the source based on the target AsClause Type of the CollectionRangeVariableSyntax.
        ''' Returns implicit BoundQueryClause or the source, in case of an early failure.
        ''' </summary>
        Private Function ApplyImplicitCollectionConversion(
                                                            syntax As CollectionRangeVariableSyntax,
                                                            source As BoundQueryPart,
                                                            variableType As TypeSymbol,
                                                            targetVariableType As TypeSymbol,
                                                            diagnostics As DiagnosticBag
                                                          ) As BoundQueryPart
            If source.Type.IsErrorType() Then
                ' If the source is already a "bad" type, we know that we will not be able to bind to the Select.
                ' Let's just report errors for the conversion between types, if any.
                Dim sourceValue As New BoundRValuePlaceholder(syntax.AsClause, variableType)
                ApplyImplicitConversion(syntax.AsClause, targetVariableType, sourceValue, diagnostics)
            Else
                ' Create LambdaSymbol for the shape of the selector.
                Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(syntax.Identifier.Identifier.ValueText, 0,
                                                                                           variableType,
                                                                                           syntax.AsClause)

                Debug.Assert(LambdaUtilities.IsNonUserCodeQueryLambda(syntax.AsClause))
                Dim lambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.ConversionNonUserCodeQueryLambda,syntax.AsClause),
                                                              ImmutableArray.Create(param))

                lambdaSymbol.SetQueryLambdaReturnType(targetVariableType)

                Dim selectorBinder As New QueryLambdaBinder(lambdaSymbol, ImmutableArray(Of RangeVariableSymbol).Empty)

                Dim selector As BoundExpression = selectorBinder.ApplyImplicitConversion(syntax.AsClause, targetVariableType,
                                                                                         New BoundParameter(param.Syntax,
                                                                                                            param,
                                                                                                            isLValue:=False,
                                                                                                            type:=variableType).MakeCompilerGenerated(),
                                                                                         diagnostics)

                Dim selectorLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                            ImmutableArray(Of RangeVariableSymbol).Empty,
                                                            selector,
                                                            exprIsOperandOfConditionalBranch:=False)
                selectorLambda.SetWasCompilerGenerated()

                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If ShouldSuppressDiagnostics(selectorLambda) Then
                    ' If the selector is already "bad", we know that we will not be able to bind to the Select.
                    ' Let's suppress additional errors.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If

                Dim boundCallOrBadExpression As BoundExpression
                boundCallOrBadExpression = BindQueryOperatorCall(syntax.AsClause, source,
                                                                 StringConstants.SelectMethod,
                                                                 ImmutableArray.Create(Of BoundExpression)(selectorLambda),
                                                                 syntax.AsClause.Span,
                                                                 diagnostics)

                Debug.Assert(boundCallOrBadExpression.WasCompilerGenerated)

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

                Return New BoundQueryClause(source.Syntax,
                                            boundCallOrBadExpression,
                                            ImmutableArray(Of RangeVariableSymbol).Empty,
                                            targetVariableType,
                                            ImmutableArray.Create(Of Binder)(selectorBinder),
                                            boundCallOrBadExpression.Type).MakeCompilerGenerated()
            End If

            Return source
        End Function

        ''' <summary>
        ''' Convert source expression to queryable type by inferring control variable type 
        ''' and applying AsQueryable/AsEnumerable or Cast(Of Object) calls.   
        ''' 
        ''' In case of success, returns possibly "converted" source and non-Nothing controlVariableType.
        ''' In case of failure, returns passed in source and Nothing as controlVariableType.
        ''' </summary>
        Private Function ConvertToQueryableType(
                                                 source As BoundExpression,
                                                 diagnostics As DiagnosticBag,
                                     <Out> ByRef controlVariableType As TypeSymbol
                                               ) As BoundExpression
            controlVariableType = Nothing

            If Not source.IsValue OrElse source.Type.IsErrorType Then Return source

            ' 11.21.2 Queryable Types
            ' A queryable collection type must satisfy one of the following conditions, in order of preference:
            ' -	It must define a conforming Select method.
            ' -	It must have one of the following methods
            ' Function AsEnumerable() As CT
            ' Function AsQueryable() As CT
            ' which can be called to obtain a queryable collection. If both methods are provided, AsQueryable is preferred over AsEnumerable.
            ' -	It must have a method
            ' Function Cast(Of T)() As CT
            ' which can be called with the type of the range variable to produce a queryable collection.

            ' Does it define a conforming Select method?
            Dim inferredType As TypeSymbol = InferControlVariableType(source, diagnostics)

            If inferredType IsNot Nothing Then
                controlVariableType = inferredType
                Return source
            End If

            Dim result As BoundExpression = Nothing
            Dim additionalDiagnostics = DiagnosticBag.GetInstance()

            ' Does it have Function AsQueryable() As CT returning queryable collection?
            Dim asQueryable As BoundExpression = BindQueryOperatorCall(source.Syntax, source, StringConstants.AsQueryableMethod,
                                                                       ImmutableArray(Of BoundExpression).Empty,
                                                                       source.Syntax.Span, additionalDiagnostics)

            If Not asQueryable.HasErrors AndAlso asQueryable.Kind = BoundKind.Call Then
                inferredType = InferControlVariableType(asQueryable, diagnostics)

                If inferredType IsNot Nothing Then
                    controlVariableType = inferredType
                    result = asQueryable
                    diagnostics.AddRange(additionalDiagnostics)
                End If
            End If

            If result Is Nothing Then
                additionalDiagnostics.Clear()

                ' Does it have Function AsEnumerable() As CT returning queryable collection?
                Dim asEnumerable As BoundExpression = BindQueryOperatorCall(source.Syntax, source, StringConstants.AsEnumerableMethod,
                                                                           ImmutableArray(Of BoundExpression).Empty,
                                                                           source.Syntax.Span, additionalDiagnostics)

                If Not asEnumerable.HasErrors AndAlso asEnumerable.Kind = BoundKind.Call Then
                    inferredType = InferControlVariableType(asEnumerable, diagnostics)

                    If inferredType IsNot Nothing Then
                        controlVariableType = inferredType
                        result = asEnumerable
                        diagnostics.AddRange(additionalDiagnostics)
                    End If
                End If
            End If

            If result Is Nothing Then
                additionalDiagnostics.Clear()

                ' If it has Function Cast(Of T)() As CT, call it with T == Object and assume Object is control variable type.
                inferredType = GetSpecialType(SpecialType.System_Object, source.Syntax, additionalDiagnostics)

                Dim cast As BoundExpression = BindQueryOperatorCall(source.Syntax, source, StringConstants.CastMethod,
                                                                    New BoundTypeArguments(source.Syntax,
                                                                                           ImmutableArray.Create(Of TypeSymbol)(inferredType)),
                                                                    ImmutableArray(Of BoundExpression).Empty,
                                                                    source.Syntax.Span, additionalDiagnostics)

                If Not cast.HasErrors AndAlso cast.Kind = BoundKind.Call Then
                    controlVariableType = inferredType
                    result = cast
                    diagnostics.AddRange(additionalDiagnostics)
                End If
            End If

            additionalDiagnostics.Free()

            Debug.Assert((result Is Nothing) = (controlVariableType Is Nothing))
            Return If(result Is Nothing, source, result)
        End Function

    End Class

End Namespace

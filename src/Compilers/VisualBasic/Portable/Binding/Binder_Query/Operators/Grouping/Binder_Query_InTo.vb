' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' This is a helper method to create a BoundQueryLambda for an Into clause 
        ''' of a Group By or a Group Join operator. 
        ''' </summary>
        Private Function BindIntoSelectorLambda(
                                                 clauseSyntax As QueryClauseSyntax,
                                                 keysRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                 keysCompoundVariableType As TypeSymbol,
                                                 addKeysInScope As Boolean,
                                                 declaredNames As HashSet(Of String),
                                                 groupType As TypeSymbol,
                                                 groupRangeVariables As ImmutableArray(Of RangeVariableSymbol),
                                                 groupCompoundVariableType As TypeSymbol,
                                                 aggregationVariables As SeparatedSyntaxList(Of AggregationRangeVariableSyntax),
                                                 mustProduceFlatCompoundVariable As Boolean,
                                                 diagnostics As DiagnosticBag,
                                     <Out> ByRef intoBinder As IntoClauseBinder,
                                     <Out> ByRef intoRangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                               ) As BoundQueryLambda
            Debug.Assert(clauseSyntax.Kind.IsKindEither(SyntaxKind.GroupByClause, SyntaxKind.GroupJoinClause))
            Debug.Assert(mustProduceFlatCompoundVariable OrElse clauseSyntax.Kind = SyntaxKind.GroupJoinClause)
            Debug.Assert((declaredNames IsNot Nothing) = (clauseSyntax.Kind = SyntaxKind.GroupJoinClause))
            Debug.Assert(keysRangeVariables.Length > 0)
            Debug.Assert(intoBinder Is Nothing)
            Debug.Assert(intoRangeVariables.IsDefault)

            Dim keyParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(keysRangeVariables), 0,
                                                                                          keysCompoundVariableType,
                                                                                          clauseSyntax, keysRangeVariables)

            Dim groupParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(StringConstants.ItAnonymous, 1,
                                                                                            groupType, clauseSyntax)

            Debug.Assert(LambdaUtilities.IsNonUserCodeQueryLambda(clauseSyntax))
            Dim intoLambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.GroupNonUserCodeQueryLambda,clauseSyntax),
                                                              ImmutableArray.Create(keyParam, groupParam))

            ' Create binder for the INTO lambda.
            Dim intoLambdaBinder As New QueryLambdaBinder(intoLambdaSymbol, ImmutableArray(Of RangeVariableSymbol).Empty)
            Dim groupReference = New BoundParameter(groupParam.Syntax, groupParam, False, groupParam.Type).MakeCompilerGenerated()

            intoBinder = New IntoClauseBinder(intoLambdaBinder,
                                              groupReference, groupRangeVariables, groupCompoundVariableType,
                                              If(addKeysInScope, keysRangeVariables.Concat(groupRangeVariables), groupRangeVariables))

            Dim intoSelector As BoundExpression = intoBinder.BindIntoSelector(clauseSyntax,
                                                                              keysRangeVariables,
                                                                              New BoundParameter(keyParam.Syntax, keyParam, False, keyParam.Type).MakeCompilerGenerated(),
                                                                              keysRangeVariables,
                                                                              Nothing,
                                                                              ImmutableArray(Of RangeVariableSymbol).Empty,
                                                                              declaredNames,
                                                                              aggregationVariables,
                                                                              mustProduceFlatCompoundVariable,
                                                                              intoRangeVariables,
                                                                              diagnostics)

            Dim intoLambda = CreateBoundQueryLambda(intoLambdaSymbol,
                                                    keysRangeVariables,
                                                    intoSelector,
                                                    exprIsOperandOfConditionalBranch:=False)

            intoLambdaSymbol.SetQueryLambdaReturnType(intoSelector.Type)
            intoLambda.SetWasCompilerGenerated()

            Return intoLambda
        End Function

    End Class

End Namespace

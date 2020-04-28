' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, if any, bind the following From operator.
        ''' 
        '''     [{Preceding query operators}] From {collection range variables}
        ''' 
        ''' Ex: From a In AA  ==> AA
        ''' 
        ''' Ex: From a In AA, b in BB  ==> AA.SelectMany(Function(a) BB, Function(a, b) New With {a, b})
        ''' 
        ''' Ex: {source with range variable 'd'} From a In AA, b in BB  ==> source.SelectMany(Function(d) AA, Function(d, a) New With {d, a}).
        '''                                                                        SelectMany(Function({d, a}) BB, 
        '''                                                                                   Function({d, a}, b) New With {d, a, b})
        ''' 
        ''' Note, that preceding Select operator can introduce unnamed range variable, which is dropped by the From
        ''' 
        ''' Ex: From a In AA Select a + 1 From b in BB ==> AA.Select(Function(a) a + 1).
        '''                                                   SelectMany(Function(unnamed) BB,
        '''                                                              Function(unnamed, b) b)  
        ''' 
        ''' Also, depending on the amount of collection range variables declared by the From, and the following query operators,
        ''' translation can produce a nested, as opposed to flat, compound variable.
        ''' 
        ''' Ex: From a In AA From b In BB, c In CC, d In DD ==> AA.SelectMany(Function(a) BB, Function(a, b) New With {a, b}).
        '''                                                        SelectMany(Function({a, b}) CC, Function({a, b}, c) New With {{a, b}, c}).
        '''                                                        SelectMany(Function({{a, b}, c}) DD, 
        '''                                                                   Function({{a, b}, c}, d) New With {a, b, c, d})   
        ''' 
        ''' If From operator translation results in a SelectMany call and the From is immediately followed by a Select or a Let operator, 
        ''' they are absorbed by the From translation. When this happens, operatorsEnumerator is advanced appropriately.
        ''' 
        ''' Ex: From a In AA From b In BB Select a + b ==> AA.SelectMany(Function(a) BB, Function(a, b) a + b)
        ''' 
        ''' Ex: From a In AA From b In BB Let c ==> AA.SelectMany(Function(a) BB, Function(a, b) new With {a, b, c})
        ''' 
        ''' </summary>
        Private Function BindFromClause(
                                         sourceOpt As BoundQueryClauseBase,
                                         from As FromClauseSyntax,
                                   ByRef operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                         diagnostics As DiagnosticBag
                                       ) As BoundQueryClauseBase
            Debug.Assert(from Is operatorsEnumerator.Current)
            Return BindCollectionRangeVariables(from, sourceOpt, from.Variables, operatorsEnumerator, diagnostics)
        End Function

        ''' <summary>
        ''' Bind query expression that starts with From keyword, as opposed to the one that starts with Aggregate.
        ''' 
        '''     From {collection range variables} [{other operators}]
        ''' </summary>
        Private Function BindFromQueryExpression(
                                                  query As QueryExpressionSyntax,
                                                  operators As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                                  diagnostics As DiagnosticBag
                                                ) As BoundQueryExpression
            ' Note, this call can advance [operators] enumerator if it absorbs the following Let or Select.
            Dim source As BoundQueryClauseBase = BindFromClause(Nothing, DirectCast(operators.Current, FromClauseSyntax), operators, diagnostics)

            source = BindSubsequentQueryOperators(source, operators, diagnostics)

            If Not source.Type.IsErrorType() AndAlso source.Kind = BoundKind.QueryableSource AndAlso
               DirectCast(source, BoundQueryableSource).Source.Kind = BoundKind.QuerySource Then
                ' Need to apply implicit Select.
                source = BindFinalImplicitSelectClause(source, diagnostics)
            End If

            Return New BoundQueryExpression(query, source, source.Type)

        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' In some scenarios, it is safe to leave compound variable in nested form when there is an
        ''' operator down the road that does its own projection (Select, Group By, ...). 
        ''' All following operators have to take an Anonymous Type in both cases and, since there is no way to
        ''' restrict the shape of the Anonymous Type in method's declaration, the operators should be
        ''' insensitive to the shape of the Anonymous Type.
        ''' </summary>
        Private Shared Function MustProduceFlatCompoundVariable(
                                                                 operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator
                                                               ) As Boolean

            While operatorsEnumerator.MoveNext()
                Select Case operatorsEnumerator.Current.Kind
                    Case SyntaxKind.SimpleJoinClause,
                         SyntaxKind.GroupJoinClause,
                         SyntaxKind.SelectClause,
                         SyntaxKind.LetClause,
                         SyntaxKind.FromClause,
                         SyntaxKind.AggregateClause
                        Return False

                    Case SyntaxKind.GroupByClause
                        ' If [Group By] doesn't have selector for a group's element, we must produce flat result. 
                        ' Element of the group can be observed through result of the query.
                        Dim groupBy = DirectCast(operatorsEnumerator.Current, GroupByClauseSyntax)
                        Return groupBy.Items.Count = 0
                End Select
            End While

            Return True
        End Function

        ''' <summary>
        ''' In some scenarios, it is safe to leave compound variable in nested form when there is an
        ''' operator down the road that does its own projection (Select, Group By, ...). 
        ''' All following operators have to take an Anonymous Type in both cases and, since there is no way to
        ''' restrict the shape of the Anonymous Type in method's declaration, the operators should be
        ''' insensitive to the shape of the Anonymous Type.
        ''' </summary>
        Private Shared Function MustProduceFlatCompoundVariable(
                                                                 groupOrInnerJoin As JoinClauseSyntax,
                                                                 operatorsEnumerator As SyntaxList(Of QueryClauseSyntax).Enumerator
                                                               ) As Boolean
            Select Case groupOrInnerJoin.Parent.Kind
                Case SyntaxKind.SimpleJoinClause
                    ' If we are nested into an Inner Join, it is safe to not flatten.
                    ' Parent join will take care of flattening.
                    Return False

                Case SyntaxKind.GroupJoinClause
                    Dim groupJoin = DirectCast(groupOrInnerJoin.Parent, GroupJoinClauseSyntax)

                    ' If we are nested into a Group Join, we are building the group for it.
                    ' It is safe to not flatten, if there is another nested join after this one,
                    ' the last nested join will take care of flattening.
                    Return groupOrInnerJoin Is groupJoin.AdditionalJoins.LastOrDefault

                Case Else
                    Return MustProduceFlatCompoundVariable(operatorsEnumerator)
            End Select
        End Function

    End Class

End Namespace

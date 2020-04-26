' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Shared Function CreateSetOfDeclaredNames() As HashSet(Of String)
            Return New HashSet(Of String)(CaseInsensitiveComparison.Comparer)
        End Function

        Private Shared Function CreateSetOfDeclaredNames(
                                                          rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                        ) As HashSet(Of String)
            Dim declaredNames As New HashSet(Of String)(CaseInsensitiveComparison.Comparer)

            For Each rangeVar As RangeVariableSymbol In rangeVariables
                declaredNames.Add(rangeVar.Name)
            Next

            Return declaredNames
        End Function

        <Conditional("DEBUG")>
        Private Shared Sub AssertDeclaredNames(
                                                declaredNames As HashSet(Of String),
                                                rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                              )
#If DEBUG Then
            For Each rangeVar As RangeVariableSymbol In rangeVariables
                If Not rangeVar.Name.StartsWith("$"c, StringComparison.Ordinal) Then
                    Debug.Assert(declaredNames.Contains(rangeVar.Name))
                End If
            Next
#End If
        End Sub

    End Class

End Namespace

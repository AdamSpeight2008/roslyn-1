' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder
        Private Shared ReadOnly s_DeclaredNamesPool As PooledObjects.ObjectPool(Of PooledObjects.PooledHashSet(Of String)) _
            = PooledObjects.PooledHashSet(Of String).CreatePool(CaseInsensitiveComparison.Comparer)

        Private Shared Function CreateSetOfDeclaredNames() As PooledHashSet(Of String)
            Return s_DeclaredNamesPool.Allocate() ' New HashSet(Of String)(CaseInsensitiveComparison.Comparer)
        End Function

        Private Shared Function CreateSetOfDeclaredNames(
                                                          rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                        ) As PooledObjects.PooledHashSet(Of String)
            Dim declaredNames = s_DeclaredNamesPool.Allocate()

            For Each rangeVar As RangeVariableSymbol In rangeVariables
                declaredNames.Add(rangeVar.Name)
            Next

            Return declaredNames
        End Function

        <Conditional("DEBUG")>
        Private Shared Sub AssertDeclaredNames(
                                                declaredNames As PooledHashSet(Of String),
                                                rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                              )
#If DEBUG Then
            Debug.Assert(declaredNames IsNot Nothing)
            For Each rangeVar As RangeVariableSymbol In rangeVariables
                If Not rangeVar.Name.StartsWith("$"c, StringComparison.Ordinal) Then
                    Debug.Assert(declaredNames.Contains(rangeVar.Name))
                End If
            Next
#End If
        End Sub

    End Class

End Namespace

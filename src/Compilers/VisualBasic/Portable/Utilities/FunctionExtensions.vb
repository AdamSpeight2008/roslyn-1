' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

'TODO - This is copied from C# and should be moved to common assemble.
Namespace Microsoft.CodeAnalysis.VisualBasic
    Friend Module HelperExts
        <Extension>
        Public Sub SwapWith(Of T)(ByRef l As T, byref r As T)
            Dim tmp = l
            l = r
            r = tmp
        End Sub
    End Module
    Friend Module FunctionExtensions
        <Extension()>
        Public Function TransitiveClosure(Of T)(relation As Func(Of T, IEnumerable(Of T)), item As T) As HashSet(Of T)
#If NETCOREAPP Then
            Dim closure = PooledObjects.PooledHashSet(Of T).GetInstance ' New HashSet(Of T)()
#Else
            Dim closure = New HashSet(Of T)()
#End If
            Dim stack as New Stack(Of T)()
            stack.Push(item)
            While stack.Count > 0
                Dim current As T = stack.Pop()
                For Each newItem In relation(current)
                    If closure.Add(newItem) Then
                        stack.Push(newItem)
                    End If
                Next
            End While
#If NETCOREAPP Then
            Dim result = closure.ToHashSet()
            closure.Free()
#Else
            Dim result = closure
#End If
            Return result
        End Function
    End Module
End Namespace

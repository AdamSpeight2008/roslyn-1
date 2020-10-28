' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Diagnostics
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic
    ''' <summary>
    '''  Structure containing all semantic information about an Await expression.
    ''' </summary>
    Public Structure AwaitExpressionInfo

        ''' <summary> Initializes a new instance of the <see cref="AwaitExpressionInfo" /> structure. </summary>
        Friend Sub New(getAwaiter As IMethodSymbol, isCompleted As IPropertySymbol, getResult As IMethodSymbol)
            Me.GetAwaiterMethod = getAwaiter
            Me.IsCompletedProperty = isCompleted
            Me.GetResultMethod = getResult
        End Sub

        ''' <summary> Gets the &quot;GetAwaiter&quot; method. </summary>
        Public ReadOnly Property GetAwaiterMethod As IMethodSymbol

        ''' <summary> Gets the &quot;GetResult&quot; method. </summary>
        Public ReadOnly Property GetResultMethod As IMethodSymbol

        ''' <summary> Gets the &quot;IsCompleted&quot; property. </summary>
        Public ReadOnly Property IsCompletedProperty As IPropertySymbol

    End Structure
End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Shared Function GetQueryLambdaParameterSyntax(
                                                               syntaxNode As VisualBasicSyntaxNode,
                                                               rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                             ) As VisualBasicSyntaxNode

            If rangeVariables.Length = 1 Then Return rangeVariables(0).Syntax

            Return syntaxNode
        End Function

        Private Function CreateQueryLambdaParameterSymbol(
                                                           name As String,
                                                           ordinal As Integer,
                                                           type As TypeSymbol,
                                                           syntaxNode As VisualBasicSyntaxNode,
                                                           rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                         ) As BoundLambdaParameterSymbol
            syntaxNode = GetQueryLambdaParameterSyntax(syntaxNode, rangeVariables)
            Dim param = New BoundLambdaParameterSymbol(name, ordinal, type, isByRef:=False, syntaxNode:=syntaxNode, location:=syntaxNode.GetLocation())
            Return param
        End Function

        Private Shared Function CreateQueryLambdaParameterSymbol(
                                                                  name As String,
                                                                  ordinal As Integer,
                                                                  type As TypeSymbol,
                                                                  syntaxNode As VisualBasicSyntaxNode
                                                                ) As BoundLambdaParameterSymbol

            Dim param = New BoundLambdaParameterSymbol(name, ordinal, type, isByRef:=False, syntaxNode:=syntaxNode, location:=syntaxNode.GetLocation())
            Return param
        End Function

        Private Shared Function GetQueryLambdaParameterName(
                                                             rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                           ) As String
            Select Case rangeVariables.Length
                   Case 0   :  Return StringConstants.ItAnonymous
                   Case 1   :  Return rangeVariables(0).Name
            End Select
            Return StringConstants.It
        End Function

        Private Shared Function GetQueryLambdaParameterNameLeft(
                                                                 rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                               ) As String
            Select Case rangeVariables.Length
                   Case 0 : Return StringConstants.ItAnonymous
                   Case 1 : Return rangeVariables(0).Name
            End Select
            Return StringConstants.It1
        End Function

        Private Shared Function GetQueryLambdaParameterNameRight(
                                                                  rangeVariables As ImmutableArray(Of RangeVariableSymbol)
                                                                ) As String
            Select Case rangeVariables.Length
                   Case 0 : Throw ExceptionUtilities.UnexpectedValue(rangeVariables.Length)
                   Case 1 : Return rangeVariables(0).Name
            End Select
            Return StringConstants.It2
        End Function

    End Class

End Namespace

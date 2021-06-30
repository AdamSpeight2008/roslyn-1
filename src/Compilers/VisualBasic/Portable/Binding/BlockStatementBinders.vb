' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Collections
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    ' The binder for a statement that can be exited with Exit and/or Continue.
    Friend Class ExitableStatementBinder
        Inherits BlockBaseBinder

        Private ReadOnly _continueLabel As LabelSymbol  ' the label to jump to for a continue (or Nothing for none)
        Private ReadOnly _continueKind As SyntaxKind    ' the kind of continue to go there
        Private ReadOnly _exitLabel As LabelSymbol      ' the label to jump to for an exit (or Nothing for none)
        Private ReadOnly _exitKind As SyntaxKind        ' the kind of exit to go there
        Private ReadOnly _loopID AS String

        Public Sub New(enclosing As Binder,
                       continueKind As SyntaxKind,
                       exitKind As SyntaxKind,
              Optional loopID As string = Nothing)
            MyBase.New(enclosing)

            Me._loopID = IF(loopID is nothing, String.Empty, "_" & loopID)

            Me._continueKind = continueKind
            If continueKind <> SyntaxKind.None Then
                Me._continueLabel = Generate_ContinueLabel()
            End If

            Me._exitKind = exitKind
            If exitKind <> SyntaxKind.None Then
                Me._exitLabel = Generate_ExitLabel()
            End If
        End Sub

        Private Function Generate_ContinueLabel() As LabelSymbol
            Dim labelName = "continue" & _loopID
            Return New GeneratedLabelSymbol(labelName)
        End Function

        Private Function Generate_ExitLabel() As LabelSymbol
            Dim labelName = "exit" & _loopID
            Return New GeneratedLabelSymbol(labelName)
        End Function


        Friend Overrides ReadOnly Property Locals As ImmutableArray(Of LocalSymbol)
            Get
                Return ImmutableArray(Of LocalSymbol).Empty
            End Get
        End Property

        Public Overrides Function GetContinueLabel(continueSyntaxKind As SyntaxKind, loopID As String) As LabelSymbol
            ' If there isn't a loopID, use existing function, that does't check the local variables.
            If loopID Is Nothing Then
                Return GetContinueLabel(continueSyntaxKind)
            End If
            ' If there is, is it contain within the locals?
            Dim found = If(loopID Is Nothing, Nothing, Locals.FirstOrDefault(Function(label) label.Name = loopID))
            If _continueKind = continueSyntaxKind AndAlso found IsNot Nothing Then
                Return _continueLabel
            Else
                ' Nope, goto next outer continuable construct.
                Return ContainingBinder.GetContinueLabel(continueSyntaxKind, loopID)
            End If
        End Function

        Public Overrides Function GetExitLabel(exitSyntaxKind As SyntaxKind, loopID As String) As LabelSymbol
            ' If there isn't a loopID, use existing function, that does't check the local variables.
            If loopID Is Nothing Then
                Return GetExitLabel(exitSyntaxKind)
            End If
            ' If there is, it is contained within the locals?
            Dim found = If(loopID Is Nothing, Nothing, Locals.FirstOrDefault(Function(label) label.Name = loopID))
            If _exitKind = exitSyntaxKind AndAlso found IsNot Nothing Then
                Return _exitLabel
            Else
                ' Nope, goto next outer exittable construct.
                Return ContainingBinder.GetExitLabel(exitSyntaxKind, loopID)
            End if
        End Function

        Public Overrides Function GetContinueLabel(continueSyntaxKind As SyntaxKind) As LabelSymbol
            If _continueKind = continueSyntaxKind Then
                Return _continueLabel
            Else
                Return ContainingBinder.GetContinueLabel(continueSyntaxKind)
            End If
        End Function

        Public Overrides Function GetExitLabel(exitSyntaxKind As SyntaxKind) As LabelSymbol
            If _exitKind = exitSyntaxKind Then
                Return _exitLabel
            Else
                Return ContainingBinder.GetExitLabel(exitSyntaxKind)
            End If
        End Function

        Public Overrides Function GetReturnLabel() As LabelSymbol
            Select Case _exitKind
                Case SyntaxKind.ExitSubStatement,
                    SyntaxKind.ExitPropertyStatement,
                    SyntaxKind.ExitFunctionStatement,
                    SyntaxKind.EventStatement,
                    SyntaxKind.OperatorStatement
                    ' the last two are not real exit statements, they are used specifically in
                    ' block events and operators to indicate that there is a return label, but
                    ' no ExitXXX statement should be able to bind to it.

                    Return _exitLabel
                Case Else
                    Return ContainingBinder.GetReturnLabel()
            End Select
        End Function
    End Class
End Namespace

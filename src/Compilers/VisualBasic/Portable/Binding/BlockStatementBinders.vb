﻿' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Collections
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    ' The binder for a statement that can be exited with Exit and/or Continue.
    Friend Class ExitableStatementBinder
        Inherits BlockBaseBinder

        Friend ReadOnly _continueLabel As LabelSymbol  ' the label to jump to for a continue (or Nothing for none)
        Friend ReadOnly _continueKind As SyntaxKind    ' the kind of continue to go there
        Friend ReadOnly _exitLabel As LabelSymbol      ' the label to jump to for an exit (or Nothing for none)
        Friend ReadOnly _exitKind As SyntaxKind        ' the kind of exit to go there
        Friend ReadOnly _ID AS String 

        Public Sub New(enclosing As Binder,
                       continueKind As SyntaxKind,
                       exitKind As SyntaxKind, Optional ID As string = Nothing)
            MyBase.New(enclosing)
            Me._ID = IF(ID is nothing,String.Empty, "_" & ID)

            Me._continueKind = continueKind
            If continueKind <> SyntaxKind.None Then
                Me._continueLabel = Generate_ContinueLabel(enclosing)
            End If

            Me._exitKind = exitKind
            If exitKind <> SyntaxKind.None Then
                Me._exitLabel = Generate_ExitLabel(enclosing)
            End If
        End Sub
        Private Function Generate_ContinueLabel(enclosing as Binder) As LabelSymbol
            Dim labelName = "continue" & _ID
            Return  New GeneratedLabelSymbol(labelName)
        End Function
        Private Function Generate_ExitLabel(enclosing as Binder) As LabelSymbol
            Dim labelName = "exit" & _ID
            Return  New GeneratedLabelSymbol(labelName)
        End Function

        Friend Overrides ReadOnly Property Locals As ImmutableArray(Of LocalSymbol)
            Get
                Return ImmutableArray(Of LocalSymbol).Empty
            End Get
        End Property

        Public Overrides Function GetContinueLabel(continueSyntaxKind As SyntaxKind) As LabelSymbol
            If _continueKind = continueSyntaxKind Then
                Return _continueLabel
            Else
                Return ContainingBinder.GetContinueLabel(continueSyntaxKind)
            End If
        End Function

        Public Overrides Function GetContinueLabel(continueSyntaxKind As SyntaxKind, controlVariable As String) As LabelSymbol
            Dim found = If(controlVariable Is Nothing, Nothing, me.Locals.FirstOrDefault(Function(i) i.Name=controlVariable))
            If  _continueKind = continueSyntaxKind AndAlso controlVariable IsNot Nothing Then
                If found IsNot Nothing Then
                Return _continueLabel
                Else
                Return ContainingBinder.GetContinueLabel(continueSyntaxKind, controlVariable)
                End If
            ElseIf _continueKind = continueSyntaxKind Then
                Return _continueLabel
            Else
                Return ContainingBinder.GetContinueLabel(continueSyntaxKind, controlVariable)
            End If
        End Function

        Public Overrides Function GetExitLabel(exitSyntaxKind As SyntaxKind, controlVariable As String) As LabelSymbol
            Dim found =  If(controlVariable Is Nothing, Nothing,me.Locals.FirstOrDefault(Function(i)  i.Name=controlVariable))
            If _exitKind = exitSyntaxKind AndAlso controlVariable IsNot Nothing Then
                If found IsNot Nothing Then
                    Return _continueLabel
                Else
                    REturn ContainingBinder.GetExitLabel(exitSyntaxKind, controlVariable)
                    End if
            Else If _exitKind = exitSyntaxKind Then
                Return _exitLabel
            Else
                Return ContainingBinder.GetExitLabel(exitSyntaxKind, controlVariable)
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

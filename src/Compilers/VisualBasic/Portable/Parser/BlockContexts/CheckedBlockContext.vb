' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

'-----------------------------------------------------------------------------
' Contains the definition of the BlockContext
'-----------------------------------------------------------------------------
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Friend NotInheritable Class CheckedBlockContext
        Inherits ExecutableStatementContext

        'Friend ReadOnly Property IsOnContext As Boolean
        'Friend ReadOnly Property IsOffContext As Boolean

        'Friend ReadOnly Property IsValid As Boolean
        '    Get
        '        Return IsOnContext Xor IsOffContext
        '    End Get
        'End Property

        Friend Sub New(statement As StatementSyntax, prevContext As BlockContext)
            MyBase.New(SyntaxKind.CheckedBlock, statement, prevContext)

            Debug.Assert(statement.Kind = SyntaxKind.BeginCheckedBlockStatement)
            'Dim s = DirectCast(statement, BeginCheckedBlockStatementSyntax)
            'IsOnContext = s.OnOrOffKeyword.Kind = SyntaxKind.OnKeyword
            'IsOffContext = s.OnOrOffKeyword.Kind = SyntaxKind.OffKeyword
        End Sub

        'Friend Overrides Function ProcessSyntax(node As VisualBasicSyntaxNode) As BlockContext
        '    If TypeOf node Is ExecutableStatementSyntax Then
        '        Add(node)
        '    Else
        '        Return MyBase.ProcessSyntax(node)
        '    End If
        '    Return Me
        'End Function

        'Friend Overrides Function TryLinkSyntax(node As VisualBasicSyntaxNode, ByRef newContext As BlockContext) As LinkResult
        '    newContext = Nothing
        '    Select Case node.Kind
        '        Case Else
        '            Return MyBase.TryLinkSyntax(node, newContext)
        '    End Select
        '    Return MyBase.TryLinkSyntax(node, newContext)
        'End Function

        Friend Overrides Function CreateBlockSyntax(endStmt As StatementSyntax) As VisualBasicSyntaxNode
            Debug.Assert(Me.BeginStatement IsNot Nothing)
            Dim beginStatement As BeginCheckedBlockStatementSyntax = DirectCast(Me.BeginStatement, BeginCheckedBlockStatementSyntax)
            If endStmt Is Nothing Then
                beginStatement = Parser.ReportSyntaxError(beginStatement, ERRID.ERR_ExpectedEndChecked)
                endStmt = SyntaxFactory.EndCheckedBlockStatement(InternalSyntaxFactory.MissingKeyword(SyntaxKind.EndKeyword),
                                                                 InternalSyntaxFactory.MissingKeyword(SyntaxKind.CheckedKeyword))
            End If
            Dim result = SyntaxFactory.CheckedBlock(beginStatement, Body(), DirectCast(endStmt, EndBlockStatementSyntax))
            FreeStatements()
            Return result
        End Function

        Friend Overrides Function EndBlock(statement As StatementSyntax) As BlockContext
            Dim blockSyntax = CreateBlockSyntax(statement)
            Return PrevBlock.ProcessSyntax(blockSyntax)
        End Function

    End Class

End Namespace

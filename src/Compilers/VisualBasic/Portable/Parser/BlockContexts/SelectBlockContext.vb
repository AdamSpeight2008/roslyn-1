' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

'-----------------------------------------------------------------------------
' Contains the definition of the BlockContext
'-----------------------------------------------------------------------------
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Friend NotInheritable Class SelectBlockContext
        Inherits ExecutableStatementContext

        Private ReadOnly _caseBlocks As SyntaxListBuilder(Of CaseBlockSyntax)
        Private _elseBlock As ElseBlockSyntax

        Friend Sub New(statement As StatementSyntax, prevContext As BlockContext)
            MyBase.New(SyntaxKind.SelectBlock, statement, prevContext)

            Debug.Assert(statement.Kind = SyntaxKind.SelectStatement)

            _caseBlocks = _parser._pool.Allocate(Of CaseBlockSyntax)()
        End Sub

        Friend Overrides Function ProcessSyntax(node As VisualBasicSyntaxNode) As BlockContext

            Select Case node.Kind
                Case SyntaxKind.CaseStatement
                    Return New CaseBlockContext(SyntaxKind.CaseBlock, DirectCast(node, StatementSyntax), Me)

                Case SyntaxKind.CaseElseStatement
                    Return New CaseBlockContext(SyntaxKind.CaseElseBlock, DirectCast(node, StatementSyntax), Me)
                Case SyntaxKind.ElseStatement
                    Return New SelectElseBlockContext(DirectCast(node, StatementSyntax), Me)
                Case SyntaxKind.CaseBlock
                    _caseBlocks.Add(DirectCast(node, CaseBlockSyntax))
                Case SyntaxKind.ElseBlock
                    _elseBlock = DirectCast(node, ElseBlockSyntax)
                Case Else
                    Return MyBase.ProcessSyntax(node)

            End Select

            Return Me
        End Function

        Friend Overrides Function TryLinkSyntax(node As VisualBasicSyntaxNode, ByRef newContext As BlockContext) As LinkResult
            newContext = Nothing

            If KindEndsBlock(node.Kind) Then
                Return UseSyntax(node, newContext)
            End If

            Select Case node.Kind

                Case _
                    SyntaxKind.CaseStatement,
                    SyntaxKind.CaseElseStatement,
                    SyntaxKind.ElseStatement
                    Return UseSyntax(node, newContext)

                ' Reuse SyntaxKind.CaseBlock but do not reuse CaseElseBlock.  These need to be crumbled so that the 
                ' error check for multiple case else statements is done.
                Case SyntaxKind.CaseBlock
                    Return UseSyntax(node, newContext) Or LinkResult.SkipTerminator

                Case SyntaxKind.ElseBlock
                    Return UseSyntax(node, newContext) Or LinkResult.SkipTerminator
                Case Else
                    ' Don't reuse other statements.  It is always an error and if a block statement is reused then the error is attached to the
                    ' block instead of the statement.
                    Return LinkResult.NotUsed
            End Select
        End Function

        Friend Overrides Function CreateBlockSyntax(endStmt As StatementSyntax) As VisualBasicSyntaxNode

            Debug.Assert(BeginStatement IsNot Nothing)
            Dim beginBlockStmt As SelectStatementSyntax = Nothing
            Dim endBlockStmt As EndBlockStatementSyntax = DirectCast(endStmt, EndBlockStatementSyntax)
            GetBeginEndStatements(beginBlockStmt, endBlockStmt)

            Dim cb= _caseBlocks.ToList
            Dim result = SyntaxFactory.SelectBlock(beginBlockStmt, cb, _elseBlock, endBlockStmt)
            _parser._pool.Free(_caseBlocks)
            FreeStatements()

            Return result
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax

'-----------------------------------------------------------------------------
' Contains the definition of the BlockContext
'-----------------------------------------------------------------------------

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Friend NotInheritable Class WhenBlockContext
        Inherits ExecutableStatementContext

        Private _whenStatement As WhenStatementSyntax
        Friend Sub New(contextKind As SyntaxKind, statement As StatementSyntax, prevContext As BlockContext)
            MyBase.New(contextKind, statement, prevContext)
            Debug.Assert((contextKind = SyntaxKind.CaseBlock AndAlso statement.Kind = SyntaxKind.WhenStatement) OrElse
                (contextKind = SyntaxKind.CaseElseBlock AndAlso statement.Kind = SyntaxKind.WhenStatement))
        End Sub

        Friend Overrides Function ProcessSyntax(node As VisualBasicSyntaxNode) As BlockContext

            Select Case node.Kind
                Case SyntaxKind.WhenStatement
                    _whenStatement = DirectCast(node, WhenStatementSyntax)
                    Return Me
                Case SyntaxKind.CaseStatement, SyntaxKind.CaseElseStatement
                    Dim context = PrevBlock.ProcessSyntax(CreateBlockSyntax(Nothing))
                    Debug.Assert(context Is PrevBlock)
                    Return context.ProcessSyntax(node)
            End Select
            Return MyBase.ProcessSyntax(node)
        End Function

        Friend Overrides Function TryLinkSyntax(node As VisualBasicSyntaxNode, ByRef newContext As BlockContext) As LinkResult
            newContext = Nothing
            Select Case node.Kind

                Case SyntaxKind.WhenStatement
                    Return UseSyntax(node, newContext)

                    'Case SyntaxKind.WhenBlock, SyntaxKind.CaseElseStatement, SyntaxKind.CaseStatement
                    '    Return LinkResult.Crumble

                Case Else
                    Return MyBase.TryLinkSyntax(node, newContext)
            End Select
        End Function

        Friend Overrides Function CreateBlockSyntax(endStmt As StatementSyntax) As VisualBasicSyntaxNode
            Debug.Assert(endStmt Is Nothing)
            Debug.Assert(BeginStatement IsNot Nothing)

            Dim result As VisualBasicSyntaxNode
            If BlockKind = SyntaxKind.WhenStatement Then
                result = SyntaxFactory.WhenBlock(DirectCast(BeginStatement, WhenStatementSyntax), Me.Body())

            ElseIf BlockKind = SyntaxKind.WhenBlock Then
                result = SyntaxFactory.WhenBlock(DirectCast(BeginStatement, WhenStatementSyntax), Me.Body())
            Else
                Throw ExceptionUtilities.UnexpectedValue(BlockKind)
            End If

            FreeStatements()

            Return result
        End Function
        Friend Overrides Function EndBlock(endStmt As StatementSyntax) As BlockContext
            Dim blockSyntax = CreateBlockSyntax(Nothing)
            Dim context = PrevBlock.ProcessSyntax(blockSyntax)
            Debug.Assert(context Is PrevBlock)
            Return PrevBlock.EndBlock(endStmt)
        End Function
    End Class

End Namespace

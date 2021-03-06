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

    Friend NotInheritable Class SelectElseBlockContext
        Inherits ExecutableStatementContext

        Friend Sub New(statement As StatementSyntax, prevContext As BlockContext)
            MyBase.New(SyntaxKind.ElseBlock, statement, prevContext)

            Debug.Assert(statement.Kind = SyntaxKind.ElseStatement)
        End Sub

        Friend Overrides Function ProcessSyntax(node As VisualBasicSyntaxNode) As BlockContext
            Debug.Assert(node IsNot Nothing)
            Select Case node.Kind
                Case SyntaxKind.ElseStatement
                    'TODO - Davidsch
                    ' In Dev10 this is reported on the keyword not the statement
                    Add(Parser.ReportSyntaxError(node, ERRID.ERR_CatchAfterFinally))
                    Return Me
            End Select

            Return MyBase.ProcessSyntax(node)
        End Function

        Friend Overrides Function KindEndsBlock(kind As SyntaxKind) As Boolean
            Return  kind = SyntaxKind.EndSelectStatement OrElse MyBase.KindEndsBlock(kind)
        End Function

        Friend Overrides Function TryLinkSyntax(node As VisualBasicSyntaxNode, ByRef newContext As BlockContext) As LinkResult
            newContext = Nothing
   
            If KindEndsBlock(node.Kind) Then
                Return UseSyntax(node, newContext)
            End If
            Select Case node.Kind

                Case SyntaxKind.ElseStatement
                    Return UseSyntax(node, newContext) Or LinkResult.Used
                Case Else
                    Return LinkResult.NotUsed    
            End Select
        End Function

        Friend Overrides Function CreateBlockSyntax(statement As StatementSyntax) As VisualBasicSyntaxNode
            Debug.Assert(statement Is Nothing)
            Debug.Assert(BeginStatement IsNot Nothing)

            Dim result = SyntaxFactory.ElseBlock(DirectCast(BeginStatement, ElseStatementSyntax), Body())

            FreeStatements()

            Return result
        End Function

        Friend Overrides Function EndBlock(statement As StatementSyntax) As BlockContext
            Dim context = PrevBlock.ProcessSyntax(CreateBlockSyntax(Nothing))
            Debug.Assert(context Is PrevBlock)
            Return context.EndBlock(statement)
        End Function

    End Class

End Namespace

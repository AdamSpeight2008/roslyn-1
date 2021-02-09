' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.Formatting.Rules
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Formatting

    Friend Class AlignTokensFormattingRule
        Inherits BaseFormattingRule
        Friend Const Name As String = "VisualBasic Align Tokens Formatting Rule"

        Public Sub New()
        End Sub

        Public Overrides Sub AddAlignTokensOperationsSlow(operations As List(Of AlignTokensOperation), node As SyntaxNode, ByRef nextOperation As NextAlignTokensOperationAction)
            nextOperation.Invoke()
            Select Case node.Kind
                   Case SyntaxKind.QueryExpression
                        AlignQueryExpressionClauses(operations, TryCast(node, QueryExpressionSyntax))
                   Case SyntaxKind.SelectBlock
                        AlignCaseBlocksToCaseInCaseBlockStatement(operations, TryCast(node, SelectBlockSyntax))
            End Select
        End Sub

        Private Shared Sub AlignQueryExpressionClauses(operations As List(Of AlignTokensOperation), queryExpression As QueryExpressionSyntax)
            ' eg
            ' Dim results = From x In xs
            '               Where x IsNot Nothing
            '               Select x
            '
            If queryExpression Is Nothing Then Return
            Dim tokens As New List(Of SyntaxToken)(queryExpression.Clauses.Select(Function(q) q.GetFirstToken(includeZeroWidth:=True)))

            If tokens.Count <= 1 Then Return
            AddAlignIndentationOfTokensToBaseTokenOperation(operations, queryExpression, tokens(0), tokens.Skip(1))
        End Sub

        Private Shared Sub AlignCaseBlocksToCaseInCaseBlockStatement(operations As List(Of AlignTokensOperation), caseBlock As SelectBlockSyntax)
            'eg.
            '  Select Case expr
            '         Case 0
            '         Case Is < 1
            '         Case Is >= 1
            '         Case Else
            '  End Select
            '            
            Dim tokens As New List(Of SyntaxToken)(caseBlock.CaseBlocks.Select(Function(q) q.GetFirstToken(includeZeroWidth:=True)))
            If tokens.Count <= 1 Then Return
            AddAlignIndentationOfTokensToBaseTokenOperation(operations, caseBlock, caseBlock.SelectStatement.CaseKeyword, tokens)
        End Sub

    End Class

End Namespace

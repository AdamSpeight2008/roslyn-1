 ' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports System.Text
Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Scanner

#Region "Comment"

        Private Function XmlMakeBeginCommentToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), scanTrailingTrivia As ScanTriviaFunc) As PunctuationSyntax
            Debug.Assert(NextAre("<!--"))
            AdvanceChar(4)
            Dim followingTrivia = scanTrailingTrivia()
            Return MakePunctuationToken(SyntaxKind.LessThanExclamationMinusMinusToken, "<!--", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeCommentToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), TokenWidth As Integer) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)

            'TODO - Normalize new lines.
            Dim text = GetTextNotInterned(TokenWidth)
            Return SyntaxFactory.XmlTextLiteralToken(text, text, precedingTrivia.Node, Nothing)

        End Function

        Private Function XmlMakeEndCommentToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            Debug.Assert(NextAre("-->"))
            AdvanceChar(3)
            Return MakePunctuationToken(SyntaxKind.MinusMinusGreaterThanToken, "-->", precedingTrivia, Nothing)
        End Function

#End Region

    End Class

End Namespace

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

        Private Function XmlMakeLeftParenToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.OpenParenToken, "(", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeRightParenToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.CloseParenToken, ")", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.EqualsToken, "=", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeDivToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.SlashToken, "/", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeColonToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.ColonToken, ":", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeGreaterToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()

            ' NOTE: > does not consume following trivia
            Return MakePunctuationToken(SyntaxKind.GreaterThanToken, ">", precedingTrivia, Nothing)
        End Function

        Private Function XmlMakeLessToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            AdvanceChar()
            Dim followingTrivia = ScanXmlWhitespace()

            Return MakePunctuationToken(SyntaxKind.LessThanToken, "<", precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeSingleQuoteToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                 spelling As Char,
                                                 isOpening As Boolean) As PunctuationSyntax
            Debug.Assert(Peek() = spelling)

            AdvanceChar()

            Dim followingTrivia As GreenNode = Nothing
            If Not isOpening Then
                Dim ws = ScanXmlWhitespace()
                followingTrivia = ws
            End If

            Return MakePunctuationToken(SyntaxKind.SingleQuoteToken, Intern(spelling), precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeDoubleQuoteToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                 spelling As Char,
                                                 isOpening As Boolean) As PunctuationSyntax
            Debug.Assert(Peek() = spelling)

            AdvanceChar()

            Dim followingTrivia As GreenNode = Nothing
            If Not isOpening Then
                Dim ws = ScanXmlWhitespace()
                followingTrivia = ws
            End If

            Return MakePunctuationToken(SyntaxKind.DoubleQuoteToken, Intern(spelling), precedingTrivia, followingTrivia)
        End Function

        Private Shared ReadOnly s_xmlAmpToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken("&amp;", "&", Nothing, Nothing)
        Private Function XmlMakeAmpLiteralToken(
                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
          ) As XmlTextTokenSyntax

            AdvanceChar(5) ' "&amp;".Length
            Return If(precedingTrivia.Node Is Nothing, s_xmlAmpToken, SyntaxFactory.XmlEntityLiteralToken("&amp;", "&", precedingTrivia.Node, Nothing))
        End Function

        Private Shared ReadOnly s_xmlAposToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken("&apos;", "'", Nothing, Nothing)
        Private Function XmlMakeAposLiteralToken(
                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
          ) As XmlTextTokenSyntax

            AdvanceChar(6) ' "&apos;".Length
            Return If(precedingTrivia.Node Is Nothing, s_xmlAposToken, SyntaxFactory.XmlEntityLiteralToken("&apos;", "'", precedingTrivia.Node, Nothing))
        End Function

        Private Shared ReadOnly s_xmlGtToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken("&gt;", ">", Nothing, Nothing)
        Private Function XmlMakeGtLiteralToken(
                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
          ) As XmlTextTokenSyntax

            AdvanceChar(4) ' "&gt;".Length
            Return If(precedingTrivia.Node Is Nothing, s_xmlGtToken, SyntaxFactory.XmlEntityLiteralToken("&gt;", "&", precedingTrivia.Node, Nothing))
        End Function

        Private Shared ReadOnly s_xmlLtToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken("&lt;", "<", Nothing, Nothing)
        Private Function XmlMakeLtLiteralToken(
                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
          ) As XmlTextTokenSyntax

            AdvanceChar(4) ' "&lt;".Length
            Return If(precedingTrivia.Node Is Nothing, s_xmlLtToken, SyntaxFactory.XmlEntityLiteralToken("&lt;", "<", precedingTrivia.Node, Nothing))
        End Function

        Private Shared ReadOnly s_xmlQuotToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken("&quot;", """", Nothing, Nothing)
        Private Function XmlMakeQuotLiteralToken(
                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
          ) As XmlTextTokenSyntax

            AdvanceChar(6) ' "&quot;".Length
            Return If(precedingTrivia.Node Is Nothing, s_xmlQuotToken, SyntaxFactory.XmlEntityLiteralToken("&quot;", """", precedingTrivia.Node, Nothing))
        End Function

    End Class

End Namespace

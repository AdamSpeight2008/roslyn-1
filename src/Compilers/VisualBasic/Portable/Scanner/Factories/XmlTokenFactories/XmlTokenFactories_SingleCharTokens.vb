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

        Private Function Make_XmlPunctuationSyntax(
                                                    text As String,
                                                    kind As SyntaxKind,
                                                    precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                           optional withTrailingTrivia As Boolean = True
                                                  ) As PunctuationSyntax
            Debug.Assert(NextAre(text))
            AdvanceChar()
            Return MakePunctuationToken(kind, text, precedingTrivia, If(withTrailingTrivia, ScanXmlWhitespace, Nothing))
        End Function

        Private Function XmlMakeLeftParenToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax("(", SyntaxKind.OpenParenToken, precedingTrivia)
        End Function

        Private Function XmlMakeRightParenToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax(")", SyntaxKind.CloseParenToken, precedingTrivia)
        End Function

        Private Function XmlMakeEqualsToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
             Return Make_XmlPunctuationSyntax("=", SyntaxKind.EqualsToken, precedingTrivia)
        End Function

        Private Function XmlMakeDivToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax("/", SyntaxKind.SlashToken, precedingTrivia)
        End Function

        Private Function XmlMakeColonToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax(":", SyntaxKind.ColonToken, precedingTrivia)
        End Function

        Private Function XmlMakeGreaterToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax(">", SyntaxKind.GreaterThanToken, precedingTrivia, withTrailingTrivia:= False)
        End Function

        Private Function XmlMakeLessToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Return Make_XmlPunctuationSyntax("<", SyntaxKind.LessThanToken, precedingTrivia)
        End Function

        Private Function Make_QuoteToken(
                                          kind as SyntaxKind,
                                          precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                          spelling As Char,
                                          isOpening As Boolean
                                        ) As PunctuationSyntax
            Debug.Assert(Peek() = spelling)

            AdvanceChar()

            Dim followingTrivia As GreenNode = Nothing
            If Not isOpening Then
                Dim ws = ScanXmlWhitespace()
                followingTrivia = ws
            End If

            Return MakePunctuationToken(kind, Intern(spelling), precedingTrivia, followingTrivia)
        End Function

        Private Function XmlMakeSingleQuoteToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  spelling As Char,
                                                  isOpening As Boolean
                                                ) As PunctuationSyntax
            Return Make_QuoteToken(SyntaxKind.SingleQuoteToken, precedingTrivia, spelling, isOpening)
        End Function

        Private Function XmlMakeDoubleQuoteToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  spelling As Char,
                                                  isOpening As Boolean
                                                ) As PunctuationSyntax
            Return Make_QuoteToken(SyntaxKind.DoubleQuoteToken, precedingTrivia, spelling, isOpening)
        End Function

#Region "&amp;"
        Private Const _amp_ = "&amp;"
        Private Shared ReadOnly s_xmlAmpToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken(_amp_, "&", Nothing, Nothing)
        Private Function XmlMakeAmpLiteralToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)  ) As XmlTextTokenSyntax
            AdvanceChar(_amp_.Length)
            Return If(precedingTrivia.Node Is Nothing, s_xmlAmpToken, WithLeadingTrivia(s_xmlAmpToken, precedingTrivia))
        End Function
#End Region
#Region "&apos;"
        Private Const _apos_ = "&apos;"
        Private Shared ReadOnly s_xmlAposToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken(_apos_, "'", Nothing, Nothing)
        Private Function XmlMakeAposLiteralToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As XmlTextTokenSyntax
            AdvanceChar(_apos_.Length)
            Return If(precedingTrivia.Node Is Nothing, s_xmlAposToken, WithLeadingTrivia(s_xmlAposToken, precedingTrivia))
        End Function
#End Region        
#Region "&gt;"
        Private Const _gt_ = "&gt;"
        Private Shared ReadOnly s_xmlGtToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken(_gt_, ">", Nothing, Nothing)
        Private Function XmlMakeGtLiteralToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As XmlTextTokenSyntax
            AdvanceChar(_gt_.Length)
            Return If(precedingTrivia.Node Is Nothing, s_xmlGtToken, WithLeadingTrivia(s_xmlGtToken, precedingTrivia))
        End Function
#End Region
#Region "&lt;"
        Private Const _lt_ = "&lt;"
        Private Shared ReadOnly s_xmlLtToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken(_lt_, "<", Nothing, Nothing)
        Private Function XmlMakeLtLiteralToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As XmlTextTokenSyntax
            AdvanceChar(_lt_.Length)
            Return If(precedingTrivia.Node Is Nothing, s_xmlLtToken, WithLeadingTrivia(s_xmlLtToken, precedingTrivia))
        End Function
#End Region
#region "&quote;"
        Private Const _quot_ = "&quot;"
        Private Shared ReadOnly s_xmlQuotToken As XmlTextTokenSyntax = SyntaxFactory.XmlEntityLiteralToken(_quot_, """", Nothing, Nothing)
        Private Function XmlMakeQuotLiteralToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As XmlTextTokenSyntax
            AdvanceChar(_quot_.Length)
            Return If(precedingTrivia.Node Is Nothing, s_xmlQuotToken, WithLeadingTrivia(s_xmlQuotToken,precedingTrivia))
        End Function
#end region

        Private Function WithLeadingTrivia(
                                            node as XmlTextTokenSyntax,
                                            precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                          ) As XmlTextTokenSyntax
            Return DirectCast(node.WithLeadingTrivia(precedingTrivia.Node), XmlTextTokenSyntax)
        End Function

    End Class
  

End Namespace

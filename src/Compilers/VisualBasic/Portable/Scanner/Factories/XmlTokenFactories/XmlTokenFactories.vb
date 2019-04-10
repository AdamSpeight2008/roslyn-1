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
        
        Private Function Make_XmlToken(
                                        text As String,
                                        kind as SyntaxKind,
                                        precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                               optional scanTrailingTrivia As ScanTriviaFunc = Nothing
                                      ) As PunctuationSyntax
            Debug.Assert(NextAre(text))
            AdvanceChar(text.Length)
            Return MakePunctuationToken(kind, text, precedingTrivia, If(scanTrailingTrivia IsNot Nothing, scanTrailingTrivia(), nothing))
        End Function

        Private Shared Function MakeMissingToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  kind As SyntaxKind
                                                ) As SyntaxToken
            Dim missing As SyntaxToken = SyntaxFactory.MissingToken(kind)
            If precedingTrivia.Any Then missing = DirectCast(missing.WithLeadingTrivia(precedingTrivia.Node), SyntaxToken)
            Return missing
        End Function

        Private Function XmlMakeBadToken(
                                          precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                          length As Integer,
                                          id As ERRID
                                        ) As BadTokenSyntax
            Return XmlMakeBadToken(SyntaxSubKind.None, precedingTrivia, length, id)
        End Function

        Private Function XmlMakeBadToken(
                                          subkind As SyntaxSubKind,
                                          precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                          length As Integer, id As ERRID
                                        ) As BadTokenSyntax

            Dim spelling = GetTextNotInterned(length)
            Dim followingTrivia = ScanXmlWhitespace()

            Dim result1 = SyntaxFactory.BadToken(subkind, spelling, precedingTrivia.Node, followingTrivia)

            Dim diagnostic As DiagnosticInfo

            Select Case id
                Case ERRID.ERR_IllegalXmlStartNameChar,
                     ERRID.ERR_IllegalXmlNameChar
                    Debug.Assert(length = 1)

                    If id = ERRID.ERR_IllegalXmlNameChar AndAlso
                        (precedingTrivia.Any OrElse
                        PrevToken Is Nothing OrElse
                        PrevToken.HasTrailingTrivia OrElse
                        PrevToken.Kind.IsEither(SyntaxKind.LessThanToken, SyntaxKind.LessThanSlashToken, SyntaxKind.LessThanQuestionToken)) Then
                        id = ERRID.ERR_IllegalXmlStartNameChar
                    End If
                    Dim xmlCh = spelling(0)
                    Dim xmlChAsUnicode = UTF16ToUnicode(New XmlCharResult(xmlCh))
                    diagnostic = ErrorFactory.ErrorInfo(id, xmlCh, String.Format("&H{0:X}", xmlChAsUnicode))
                Case Else
                    diagnostic = ErrorFactory.ErrorInfo(id)
            End Select

            Dim errResult1 = DirectCast(result1.SetDiagnostics({diagnostic}), BadTokenSyntax)
            Debug.Assert(errResult1 IsNot Nothing)
            Return errResult1
        End Function

        Private Function XmlMakeXmlNCNameToken(
                                                precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                TokenWidth As Integer
                                              ) As XmlNameTokenSyntax
            Debug.Assert(TokenWidth > 0)
            Dim text = GetText(TokenWidth)
            'Xml/Version/Standalone/Encoding/DOCTYPE
            Dim contextualKind As SyntaxKind = SyntaxKind.XmlNameToken
            Select Case text.Length
                Case 3  : If String.Equals(text, "xml", StringComparison.Ordinal) Then contextualKind = SyntaxKind.XmlKeyword
            End Select

            If contextualKind = SyntaxKind.XmlNameToken Then
                contextualKind = TokenOfStringCached(text)
                If contextualKind = SyntaxKind.IdentifierToken Then  contextualKind = SyntaxKind.XmlNameToken
            End If
            Dim followingTrivia = ScanXmlWhitespace()
            Return SyntaxFactory.XmlNameToken(text, contextualKind, precedingTrivia.Node, followingTrivia)
        End Function

        Private Function XmlMakeAttributeDataToken(
                                                    precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                    TokenWidth As Integer,
                                                    Value As String
                                                  ) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)
            Dim text = GetTextNotInterned(TokenWidth)
            ' NOTE: XmlMakeAttributeData does not consume trailing trivia.
            Return SyntaxFactory.XmlTextLiteralToken(text, Value, precedingTrivia.Node, Nothing)
        End Function

        Private Function XmlMakeAttributeDataToken(
                                                    precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                    TokenWidth As Integer,
                                                    Scratch As StringBuilder
                                                  ) As XmlTextTokenSyntax
            ' NOTE: XmlMakeAttributeData does not consume trailing trivia.
            Return XmlMakeTextLiteralToken(precedingTrivia, TokenWidth, Scratch)
        End Function

        Private Function XmlMakeEntityLiteralToken(
                                                    precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                    TokenWidth As Integer,
                                                    Value As String
                                                  ) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)
            Return SyntaxFactory.XmlEntityLiteralToken(GetText(TokenWidth), Value, precedingTrivia.Node, Nothing)
        End Function

        Private Function XmlMakeTextLiteralToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  TokenWidth As Integer,
                                                  Scratch As StringBuilder
                                                ) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)
            Dim text = GetTextNotInterned(TokenWidth)
            ' PERF: It's common for the text and the 'value' to be identical. If so, try to unify the
            ' two strings.
            Dim value = GetScratchText(Scratch, text)
            Return SyntaxFactory.XmlTextLiteralToken(text, value, precedingTrivia.Node, Nothing)
        End Function

        Private Shared ReadOnly s_docCommentCrLfToken As XmlTextTokenSyntax = SyntaxFactory.DocumentationCommentLineBreakToken(vbCrLf, vbLf, Nothing, Nothing)

        Private Function MakeDocCommentLineBreakToken(
                                                       precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                       TokenWidth As Integer
                                                     ) As XmlTextTokenSyntax
            Dim text = GetText(TokenWidth)
            Debug.Assert(text = vbCr OrElse text = vbLf OrElse text = vbCrLf)
            If precedingTrivia.Node Is Nothing AndAlso text = vbCrLf Then Return s_docCommentCrLfToken
            Return SyntaxFactory.DocumentationCommentLineBreakToken(text, vbLf, precedingTrivia.Node, Nothing)
        End Function

        Private Function XmlMakeTextLiteralToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  TokenWidth As Integer,
                                                  err As ERRID
                                                ) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)
            Dim text = GetTextNotInterned(TokenWidth)
            Return DirectCast(SyntaxFactory.XmlTextLiteralToken(text, text, precedingTrivia.Node, Nothing).SetDiagnostics({ErrorFactory.ErrorInfo(err)}), XmlTextTokenSyntax)
        End Function

        Private Function XmlMakeBeginEndElementToken(
                                                      precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), 
                                                      scanTrailingTrivia As ScanTriviaFunc
                                                    ) As PunctuationSyntax
            Return Make_XmlToken("</", SyntaxKind.LessThanSlashToken, precedingTrivia, scanTrailingTrivia)
        End Function

        Private Function XmlMakeEndEmptyElementToken(
                                                      precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                                    ) As PunctuationSyntax
            Return Make_XmlToken("/>", SyntaxKind.SlashGreaterThanToken, precedingTrivia)
        End Function

    End Class

End Namespace

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax
    Partial Friend Class Scanner

#Region "Punctuation"
        Friend Function MakePunctuationToken(
                                              precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                              spelling As String,
                                              kind As SyntaxKind
                                            ) As PunctuationSyntax
            Dim followingTrivia = ScanSingleLineTrivia()
            Return MakePunctuationToken(kind, spelling, precedingTrivia, followingTrivia)
        End Function

        Private Function MakePunctuationToken(
                                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                               length As Integer,
                                               kind As SyntaxKind
                                             ) As PunctuationSyntax
            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()
            Return MakePunctuationToken(kind, spelling, precedingTrivia, followingTrivia)
        End Function

        Friend Function MakePunctuationToken(
                                              kind As SyntaxKind,
                                              spelling As String,
                                              precedingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode),
                                              followingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode)
                                            ) As PunctuationSyntax
            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)
            Dim p As PunctuationSyntax = Nothing
            If _punctTable.TryGetValue(tp, p) Then Return p
            p = New PunctuationSyntax(kind, spelling, precedingTrivia.Node, followingTrivia.Node)
            _punctTable.Add(tp, p)
            Return p
        End Function

        Friend Function MakePunctuationToken(
                                               charIsFullWidth As Boolean,
                                               ch As Char,
                                               fw As Char,
                                               kind As SyntaxKind,
                                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode)
                                            ) As PunctuationSyntax
            Dim spelling As String = If(charIsFullWidth, fw, ch)
            AdvanceChar()
            Return MakePunctuationToken(precedingTrivia, spelling, kind)
        End Function

        Private Function MakeOpenParenToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "("c, FULLWIDTH_LEFT_PARENTHESIS, SyntaxKind.OpenParenToken, precedingTrivia)
        End Function

        Private Function MakeCloseParenToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, ")"c, FULLWIDTH_RIGHT_PARENTHESIS, SyntaxKind.CloseParenToken, precedingTrivia)
        End Function

        Private Function MakeDotToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth,  "."c, FULLWIDTH_FULL_STOP, SyntaxKind.DotToken, precedingTrivia)
        End Function

        Private Function MakeCommaToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth,  ","c, FULLWIDTH_COMMA, SyntaxKind.CommaToken, precedingTrivia)
        End Function

        Private Function MakeEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "="c, FULLWIDTH_EQUALS_SIGN, SyntaxKind.EqualsToken, precedingTrivia)
        End Function

        Private Function MakeHashToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth,  "#"c, FULLWIDTH_NUMBER_SIGN, SyntaxKind.HashToken, precedingTrivia)
        End Function

        Private Function MakeAmpersandToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth,  "&"c, FULLWIDTH_AMPERSAND, SyntaxKind.AmpersandToken, precedingTrivia)
        End Function

        Private Function MakeOpenBraceToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "{"c, FULLWIDTH_LEFT_CURLY_BRACKET, SyntaxKind.OpenBraceToken, precedingTrivia)
        End Function

        Private Function MakeCloseBraceToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "}"c, FULLWIDTH_RIGHT_CURLY_BRACKET, SyntaxKind.CloseBraceToken, precedingTrivia)
        End Function

        Private Function MakeColonToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Debug.Assert(Peek() = If(charIsFullWidth, FULLWIDTH_COLON, ":"c))
            Debug.Assert(Not precedingTrivia.Any())

            Dim width = _endOfTerminatorTrivia - _lineBufferOffset
            Debug.Assert(width = 1)

            AdvanceChar(width)

            ' Colon does not consume trailing trivia
            Return SyntaxFactory.ColonToken
        End Function

        Private Function MakeEmptyToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, "", SyntaxKind.EmptyToken)
        End Function

        Private Function MakePlusToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "+"c, FULLWIDTH_PLUS_SIGN, SyntaxKind.PlusToken, precedingTrivia)
        End Function

        Private Function MakeMinusToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "-"c, FULLWIDTH_HYPHEN_MINUS, SyntaxKind.MinusToken, precedingTrivia)
        End Function

        Private Function MakeAsteriskToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "*"c, FULLWIDTH_ASTERISK, SyntaxKind.AsteriskToken, precedingTrivia)
        End Function

        Private Function MakeSlashToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "/"c, FULLWIDTH_SOLIDUS, SyntaxKind.SlashToken, precedingTrivia)
        End Function

        Private Function MakeBackslashToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "\"c, FULLWIDTH_REVERSE_SOLIDUS, SyntaxKind.BackslashToken, precedingTrivia)
        End Function

        Private Function MakeCaretToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "^"c, FULLWIDTH_CIRCUMFLEX_ACCENT, SyntaxKind.CaretToken, precedingTrivia)
        End Function

        Private Function MakeExclamationToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, "!"c, FULLWIDTH_EXCLAMATION_MARK, SyntaxKind.ExclamationToken, precedingTrivia)
        End Function

        Private Function MakeQuestionToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
              Return MakePunctuationToken(charIsFullWidth, "?"c, FULLWIDTH_QUESTION_MARK, SyntaxKind.QuestionToken, precedingTrivia)
        End Function

        Private Function MakeGreaterThanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
             Return MakePunctuationToken(charIsFullWidth, ">"c, FULLWIDTH_GREATER_THAN_SIGN, SyntaxKind.GreaterThanToken, precedingTrivia)
        End Function

        Private Function MakeLessThanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
              Return MakePunctuationToken(charIsFullWidth, "<"c, FULLWIDTH_LESS_THAN_SIGN, SyntaxKind.LessThanToken, precedingTrivia)
        End Function

        ' ==== TOKENS WITH NOT FIXED SPELLING

        Private Function MakeStatementTerminatorToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), width As Integer) As PunctuationSyntax
            Debug.Assert(_endOfTerminatorTrivia = _lineBufferOffset + width)
            Debug.Assert(width = 1 OrElse width = 2)
            Debug.Assert(Not precedingTrivia.Any())

            AdvanceChar(width)

            ' Statement terminator does not consume trailing trivia
            Return SyntaxFactory.StatementTerminatorToken
        End Function

        Private Function MakeAmpersandEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.AmpersandEqualsToken)
        End Function

        Private Function MakeColonEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.ColonEqualsToken)
        End Function

        Private Function MakePlusEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.PlusEqualsToken)
        End Function

        Private Function MakeMinusEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.MinusEqualsToken)
        End Function

        Private Function MakeAsteriskEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.AsteriskEqualsToken)
        End Function

        Private Function MakeSlashEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.SlashEqualsToken)
        End Function

        Private Function MakeBackSlashEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.BackslashEqualsToken)
        End Function

        Private Function MakeCaretEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.CaretEqualsToken)
        End Function

        Private Function MakeGreaterThanEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.GreaterThanEqualsToken)
        End Function

        Private Function MakeLessThanEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.LessThanEqualsToken)
        End Function

        Private Function MakeLessThanGreaterThanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.LessThanGreaterThanToken)
        End Function

        Private Function MakeLessThanLessThanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.LessThanLessThanToken)
        End Function

        Private Function MakeGreaterThanGreaterThanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.GreaterThanGreaterThanToken)
        End Function

        Private Function MakeLessThanLessThanEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.LessThanLessThanEqualsToken)
        End Function

        Private Function MakeGreaterThanGreaterThanEqualsToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), length As Integer) As PunctuationSyntax
            Return MakePunctuationToken(precedingTrivia, length, SyntaxKind.GreaterThanGreaterThanEqualsToken)
        End Function

        Private Function MakeAtToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As PunctuationSyntax
            Return MakePunctuationToken(charIsFullWidth, "@"c, FULLWIDTH_COMMERCIAL_AT, SyntaxKind.AtToken, precedingTrivia)
        End Function

#End Region

    End Class

End Namespace

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Strict On
Option Explicit On
Option Infer On

Imports System.Globalization

Namespace Microsoft.CodeAnalysis.VisualBasic
    
    ''' <summary>
    ''' Provides members for determining Syntax facts about characters and Unicode conversions.
    ''' </summary>
    Partial Public Class SyntaxFacts
        Private Const s_fullwidth = CInt(&HFF00L - &H0020L)

        '/*****************************************************************************/
        '// MakeFullWidth - Converts a half-width to full-width character
        Friend Shared Function MakeFullWidth(c As Char) As Char
            Debug.Assert(IsHalfWidth(c))
            Return Convert.ToChar(Convert.ToUInt16(c) + s_fullwidth)
        End Function

        Friend Shared Function IsHalfWidth(c As Char) As Boolean
            Return c >= ChrW(&H21S) AndAlso c <= ChrW(&H7ES)
        End Function

        '// MakeHalfWidth - Converts a full-width character to half-width
        Friend Shared Function MakeHalfWidth(c As Char) As Char
            Debug.Assert(IsFullWidth(c))

            Return Convert.ToChar(Convert.ToUInt16(c) - s_fullwidth)
        End Function

        '// IsFullWidth - Returns if the character is full width
        Friend Shared Function IsFullWidth(c As Char) As Boolean
            ' Do not use "AndAlso" or it will not inline.
            Return c > ChrW(&HFF00US) And c < ChrW(&HFF5FUS)
        End Function

        ''' <summary>
        ''' Determines if Unicode character represents a whitespace.
        ''' </summary>
        ''' <param name="c">The Unicode character.</param>
        ''' <returns>A boolean value set to True if character represents whitespace.</returns>
        Public Shared Function IsWhitespace(c As Char) As Boolean
            Return (SPACE = c) OrElse (CHARACTER_TABULATION = c) OrElse (c > ChrW(128) AndAlso IsWhitespaceNotAscii(c))
        End Function

        ''' <summary>
        ''' Determines if Unicode character represents a XML whitespace.
        ''' </summary>
        ''' <param name="c">The unicode character</param>
        ''' <returns>A boolean value set to True if character represents XML whitespace.</returns>
        Public Shared Function IsXmlWhitespace(c As Char) As Boolean
            Return (SPACE = c) OrElse (CHARACTER_TABULATION = c) OrElse (c > ChrW(128) AndAlso XmlCharType.IsWhiteSpace(c))
        End Function

        Friend Shared Function IsWhitespaceNotAscii(ch As Char) As Boolean
            Select Case ch
                Case NO_BREAK_SPACE, IDEOGRAPHIC_SPACE, ChrW(&H2000S) To ChrW(&H200BS)
                    Return True

                Case Else
                    Return CharUnicodeInfo.GetUnicodeCategory(ch) = UnicodeCategory.SpaceSeparator
            End Select
        End Function




        Friend Const CHARACTER_TABULATION               = ChrW(&H0009)
        Friend Const LINE_FEED                          = ChrW(&H000A)
        Friend Const CARRIAGE_RETURN                    = ChrW(&H000D)
        Friend Const SPACE                              = ChrW(&H0020)
        Friend Const NO_BREAK_SPACE                     = ChrW(&H00A0)
        Friend Const IDEOGRAPHIC_SPACE                  = ChrW(&H3000)
        Friend Const LINE_SEPARATOR                     = ChrW(&H2028)
        Friend Const PARAGRAPH_SEPARATOR                = ChrW(&H2029)
        Friend Const NEXT_LINE                          = ChrW(&H0085)


        Friend Const  LEFT_SINGLE_QUOTATION_MARK        = ChrW(&H2018)  REM ‘
        Friend Const RIGHT_SINGLE_QUOTATION_MARK        = ChrW(&H2019)  REM ’

        Friend Const  LEFT_DOUBLE_QUOTATION_MARK        = ChrW(&H201C)  REM “
        Friend Const RIGHT_DOUBLE_QUOTATION_MARK        = ChrW(&H201D)  REM ”

        Friend Const FULLWIDTH_APOSTROPHE               = ChrW(s_fullwidth + AscW("'"c))  REM ＇
        Friend Const FULLWIDTH_QUOTATION_MARK           = ChrW(s_fullwidth + AscW(""""c)) REM ＂

        Friend Const FULLWIDTH_DIGIT_ZERO               = ChrW(s_fullwidth + AscW("0"c))  REM ０
        Friend Const FULLWIDTH_DIGIT_ONE                = ChrW(s_fullwidth + AscW("1"c))  REM １
        Friend Const FULLWIDTH_DIGIT_SEVEN              = ChrW(s_fullwidth + AscW("7"c))  REM ７
        Friend Const FULLWIDTH_DIGIT_NINE               = ChrW(s_fullwidth + AscW("9"c))  REM ９

        Friend Const FULLWIDTH_LOW_LINE                 = ChrW(s_fullwidth + AscW("_"c))  REM ＿
        Friend Const FULLWIDTH_COLON                    = ChrW(s_fullwidth + AscW(":"c))  REM ：
        Friend Const FULLWIDTH_SOLIDUS                  = ChrW(s_fullwidth + AscW("/"c))  REM ／
        Friend Const FULLWIDTH_HYPHEN_MINUS             = ChrW(s_fullwidth + AscW("-"c))  REM －
        Friend Const FULLWIDTH_PLUS_SIGN                = ChrW(s_fullwidth + AscW("+"c))  REM ＋
        Friend Const FULLWIDTH_NUMBER_SIGN              = ChrW(s_fullwidth + AscW("#"c))  REM ＃

        Friend Const FULLWIDTH_EQUALS_SIGN              = ChrW(s_fullwidth + AscW("="c))  REM ＝
        Friend Const FULLWIDTH_LESS_THAN_SIGN           = ChrW(s_fullwidth + AscW("<"c))  REM ＜
        Friend Const FULLWIDTH_GREATER_THAN_SIGN        = ChrW(s_fullwidth + AscW(">"c))  REM ＞
        Friend Const FULLWIDTH_LEFT_PARENTHESIS         = ChrW(s_fullwidth + AscW("("c))  REM （
        Friend Const FULLWIDTH_RIGHT_PARENTHESIS        = ChrW(s_fullwidth + AscW(")"c))
        Friend Const FULLWIDTH_LEFT_SQUARE_BRACKET      = ChrW(s_fullwidth + AscW("["c))  REM ［
        Friend Const FULLWIDTH_RIGHT_SQUARE_BRACKET     = ChrW(s_fullwidth + AscW("]"c))  REM ］
        Friend Const FULLWIDTH_LEFT_CURLY_BRACKET       = ChrW(s_fullwidth + AscW("{"c))  REM ｛
        Friend Const FULLWIDTH_RIGHT_CURLY_BRACKET      = ChrW(s_fullwidth + AscW("}"c))  REM ｝
        Friend Const FULLWIDTH_AMPERSAND                = ChrW(s_fullwidth + AscW("&"c))  REM ＆
        Friend Const FULLWIDTH_DOLLAR_SIGN              = ChrW(s_fullwidth + AscW("$"c))  REM ＄
        Friend Const FULLWIDTH_QUESTION_MARK            = ChrW(s_fullwidth + AscW("?"c))  REM ？
        Friend Const FULLWIDTH_FULL_STOP                = ChrW(s_fullwidth + AscW("."c))  REM ．
        Friend Const FULLWIDTH_COMMA                    = ChrW(s_fullwidth + AscW(","c))  REM ，
        Friend Const FULLWIDTH_PERCENT_SIGN             = ChrW(s_fullwidth + AscW("%"c))  REM ％
        Friend Const FULLWIDTH_ASTERISK                 = ChrW(s_fullwidth + AscW("*"c))
        Friend Const FULLWIDTH_REVERSE_SOLIDUS          = ChrW(s_fullwidth + AscW("\"c))
        Friend Const FULLWIDTH_CIRCUMFLEX_ACCENT        = ChrW(s_fullwidth + AscW("^"c))
        Friend Const FULLWIDTH_EXCLAMATION_MARK         = ChrW(s_fullwidth + AscW("!"c))
        Friend Const FULLWIDTH_COMMERCIAL_AT            = ChrW(s_fullwidth + AscW("@"c))

        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_B   = ChrW(s_fullwidth + AscW("B"c))  REM Ｂ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_H   = ChrW(s_fullwidth + AscW("H"c))  REM Ｈ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_O   = ChrW(s_fullwidth + AscW("O"c))  REM Ｏ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_E   = ChrW(s_fullwidth + AscW("E"c))  REM Ｅ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_A   = ChrW(s_fullwidth + AscW("A"c))  REM Ａ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_F   = ChrW(s_fullwidth + AscW("F"c))  REM Ｆ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_C   = ChrW(s_fullwidth + AscW("C"c))  REM Ｃ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_P   = ChrW(s_fullwidth + AscW("P"c))  REM Ｐ
        Friend Const FULLWIDTH_LATIN_CAPITAL_LETTER_M   = ChrW(s_fullwidth + AscW("M"c))  REM Ｍ

        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_B     = ChrW(s_fullwidth + AscW("b"c))  REM ｂ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_H     = ChrW(s_fullwidth + AscW("h"c))  REM ｈ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_O     = ChrW(s_fullwidth + AscW("o"c))  REM ｏ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_E     = ChrW(s_fullwidth + AscW("e"c))  REM ｅ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_A     = ChrW(s_fullwidth + AscW("a"c))  REM ａ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_F     = ChrW(s_fullwidth + AscW("f"c))  REM ｆ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_C     = ChrW(s_fullwidth + AscW("c"c))  REM ｃ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_P     = ChrW(s_fullwidth + AscW("p"c))  REM ｐ
        Friend Const FULLWIDTH_LATIN_SMALL_LETTER_M     = ChrW(s_fullwidth + AscW("m"c))  REM ｍ

        ''' <summary>
        ''' Determines if the Unicode character is a newline character.
        ''' </summary>
        ''' <param name="c">The Unicode character.</param>
        ''' <returns>A boolean value set to True if character is a newline character.</returns>
        Public Shared Function IsNewLine(c As Char) As Boolean
            Return CARRIAGE_RETURN = c OrElse LINE_FEED = c OrElse (c >= NEXT_LINE AndAlso (NEXT_LINE = c OrElse LINE_SEPARATOR = c OrElse PARAGRAPH_SEPARATOR = c))
        End Function

        Friend Shared Function IsSingleQuote(c As Char) As Boolean
            ' // Besides the half width and full width ', we also check for Unicode
            ' // LEFT SINGLE QUOTATION MARK and RIGHT SINGLE QUOTATION MARK because
            ' // IME editors paste them in. This isn't really technically correct
            ' // because we ignore the left-ness or right-ness, but see VS 170991
            Return c = "'"c OrElse (c >= LEFT_SINGLE_QUOTATION_MARK AndAlso (c = FULLWIDTH_APOSTROPHE Or c = LEFT_SINGLE_QUOTATION_MARK Or c = RIGHT_SINGLE_QUOTATION_MARK))
        End Function

        Friend Shared Function IsDoubleQuote(c As Char) As Boolean
            ' // Besides the half width and full width ", we also check for Unicode
            ' // LEFT DOUBLE QUOTATION MARK and RIGHT DOUBLE QUOTATION MARK because
            ' // IME editors paste them in. This isn't really technically correct
            ' // because we ignore the left-ness or right-ness, but see VS 170991
            Return c = """"c OrElse (c >= LEFT_DOUBLE_QUOTATION_MARK AndAlso (c = FULLWIDTH_QUOTATION_MARK Or c = LEFT_DOUBLE_QUOTATION_MARK Or c = RIGHT_DOUBLE_QUOTATION_MARK))
        End Function

        Friend Shared Function IsLeftCurlyBracket(c As Char) As Boolean
            Return c = "{"c OrElse c = FULLWIDTH_LEFT_CURLY_BRACKET
        End Function

        Friend Shared Function IsRightCurlyBracket(c As Char) As Boolean
            Return c = "}"c OrElse c = FULLWIDTH_RIGHT_CURLY_BRACKET
        End Function

        ''' <summary>
        ''' Determines if the unicode character is a colon character.
        ''' </summary>
        ''' <param name="c">The unicode character.</param>
        ''' <returns>A boolean value set to True if character is a colon character.</returns>
        Public Shared Function IsColon(c As Char) As Boolean
            Return c = ":"c OrElse c = FULLWIDTH_COLON
        End Function

        ''' <summary>
        ''' Determines if the unicode character is a underscore character.
        ''' </summary>
        ''' <param name="c">The unicode character.</param>
        ''' <returns>A boolean value set to True if character is an underscore character.</returns>
        Public Shared Function IsUnderscore(c As Char) As Boolean
            ' NOTE: fullwidth _ is not considered any special.
            Return c = "_"c
        End Function

        ''' <summary>
        ''' Determines if the unicode character is a hash character.
        ''' </summary>
        ''' <param name="c">The unicode character.</param>
        ''' <returns>A boolean value set to True if character is a hash character.</returns>
        Public Shared Function IsHash(c As Char) As Boolean
            Return c = "#"c OrElse c = FULLWIDTH_NUMBER_SIGN
        End Function

        ''' <summary>
        ''' Determines if the Unicode character can be the starting character of a Visual Basic identifier.
        ''' </summary>
        ''' <param name="c">The Unicode character.</param>
        ''' <returns>A boolean value set to True if character can be part of a valid start character in an identifier.</returns>
        Public Shared Function IsIdentifierStartCharacter(
                                                           c As Char
                                                         ) As Boolean
            'TODO: make easy cases fast (or check if they already are)
            Dim CharacterProperties As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c)

            Return IsPropAlpha(CharacterProperties) OrElse
            IsPropLetterDigit(CharacterProperties) OrElse
            IsPropConnectorPunctuation(CharacterProperties)
        End Function

        ' TODO: replace CByte with something faster.
        Friend Shared Function IntegralLiteralCharacterValue(
                                                              Digit As Char
                                                            ) As Byte
            If IsFullWidth(Digit) Then Digit = MakeHalfWidth(Digit)
            Dim u As Integer = AscW(Digit)

            If IsDecimalDigit(Digit)               Then Return CByte(u - AscW("0"c))
            If Digit >= "A"c AndAlso Digit <= "F"c Then Return CByte(u + (10 - AscW("A"c)))
            Debug.Assert(Digit >= "a"c AndAlso Digit <= "f"c, "Surprising digit.")
            Return CByte(u + (10 - AscW("a"c)))
          End Function

        Friend Shared Function BeginsBaseLiteral(c As Char) As Boolean
            Return (c = "H"c Or c = "O"c Or c = "B"c Or c = "h"c Or c = "o"c Or c = "b"c) OrElse
                    (IsFullWidth(c) AndAlso (c = FULLWIDTH_LATIN_CAPITAL_LETTER_H Or c = FULLWIDTH_LATIN_SMALL_LETTER_H) Or
                                            (c = FULLWIDTH_LATIN_CAPITAL_LETTER_O Or c = FULLWIDTH_LATIN_SMALL_LETTER_O) Or
                                            (c = FULLWIDTH_LATIN_CAPITAL_LETTER_B Or c = FULLWIDTH_LATIN_SMALL_LETTER_B))
        End Function


        Private Shared ReadOnly s_isIDChar As Boolean() =
        {
            False, False, False, False, False, False, False, False, False, False,
            False, False, False, False, False, False, False, False, False, False,
            False, False, False, False, False, False, False, False, False, False,
            False, False, False, False, False, False, False, False, False, False,
            False, False, False, False, False, False, False, False, True, True,
            True, True, True, True, True, True, True, True, False, False,
            False, False, False, False, False, True, True, True, True, True,
            True, True, True, True, True, True, True, True, True, True,
            True, True, True, True, True, True, True, True, True, True,
            True, False, False, False, False, True, False, True, True, True,
            True, True, True, True, True, True, True, True, True, True,
            True, True, True, True, True, True, True, True, True, True,
            True, True, True, False, False, False, False, False
        }

        Friend Shared Function IsNarrowIdentifierCharacter(c As UInt16) As Boolean
            Return s_isIDChar(c)
        End Function

        ''' <summary>
        ''' Determines if the Unicode character can be a part of a Visual Basic identifier.
        ''' </summary>
        ''' <param name="c">The Unicode character.</param>
        ''' <returns>A boolean value set to True if character can be part of a valid identifier.</returns>
        Public Shared Function IsIdentifierPartCharacter(c As Char) As Boolean
            If c < ChrW(128) Then Return IsNarrowIdentifierCharacter(Convert.ToUInt16(c))
            Return IsWideIdentifierCharacter(c)
        End Function

        ''' <summary>
        ''' Determines if the name is a valid identifier.
        ''' </summary>
        ''' <param name="name">The identifier name.</param>
        ''' <returns>A boolean value set to True if name is valid identifier.</returns>
        Public Shared Function IsValidIdentifier(
                                                  name As String
                                                ) As Boolean
            If String.IsNullOrEmpty(name) orelse Not IsIdentifierStartCharacter(name(0)) Then Return False
            Dim nameLength As Integer = name.Length
            For i As Integer = 1 To nameLength - 1 ' NB: start at 1
                If Not IsIdentifierPartCharacter(name(i)) Then Return False
            Next

            Return True
        End Function

        ''' <summary>
        ''' Creates a half width form Unicode character string. 
        ''' </summary>
        ''' <param name="text">The text representing the original identifier.  This can be in full width or half width Unicode form.  </param>
        ''' <returns>A string representing the text in a half width Unicode form.</returns>
        Public Shared Function MakeHalfWidthIdentifier(
                                                        text As String
                                                      ) As String
            If text Is Nothing Then Return text

            Dim characters As Char() = Nothing
            For i = 0 To text.Length - 1
                Dim c = text(i)

                If IsFullWidth(c) Then
                    If characters Is Nothing Then
                        characters = New Char(text.Length - 1) {}
                        text.CopyTo(0, characters, 0, i)
                    End If

                    characters(i) = MakeHalfWidth(c)
                ElseIf characters IsNot Nothing Then
                    characters(i) = c
                End If
            Next

            Return If(characters Is Nothing, text, New String(characters))
        End Function

        Friend Shared Function IsWideIdentifierCharacter(c As Char) As Boolean
            Dim CharacterProperties As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c)

            Return IsPropAlphaNumeric(CharacterProperties)          OrElse IsPropLetterDigit(CharacterProperties) OrElse
                   IsPropConnectorPunctuation(CharacterProperties)  OrElse IsPropCombining(CharacterProperties)   OrElse
                   IsPropOtherFormat(CharacterProperties)
        End Function

        Friend Shared Function BeginsExponent(c As Char) As Boolean
            Return c = "E"c Or c = "e"c Or c = FULLWIDTH_LATIN_CAPITAL_LETTER_E Or c = FULLWIDTH_LATIN_SMALL_LETTER_E
        End Function

        Friend Shared Function IsBinaryDigit(c As Char) As Boolean
            Return (c >= "0"c And c <= "1"c) Or
                   (c >= FULLWIDTH_DIGIT_ZERO And c <= FULLWIDTH_DIGIT_ONE)
        End Function

        Friend Shared Function IsOctalDigit(c As Char) As Boolean
            Return (c >= "0"c And c <= "7"c) Or
                   (c >= FULLWIDTH_DIGIT_ZERO And c <= FULLWIDTH_DIGIT_SEVEN)
        End Function

        Friend Shared Function IsDecimalDigit(c As Char) As Boolean
            Return (c >= "0"c And c <= "9"c) Or
                   (c >= FULLWIDTH_DIGIT_ZERO And c <= FULLWIDTH_DIGIT_NINE)
        End Function

        Friend Shared Function IsHexDigit(c As Char) As Boolean
            Return IsDecimalDigit(c) OrElse
                    (c >= "a"c And c <= "f"c) OrElse
                    (c >= "A"c And c <= "F"c) OrElse
                    (c >= FULLWIDTH_LATIN_SMALL_LETTER_A And c <= FULLWIDTH_LATIN_SMALL_LETTER_F) OrElse
                    (c >= FULLWIDTH_LATIN_CAPITAL_LETTER_A And c <= FULLWIDTH_LATIN_CAPITAL_LETTER_F)
        End Function

        Friend Shared Function IsDateSeparatorCharacter(c As Char) As Boolean
            Return c = "/"c Or c = "-"c Or c = FULLWIDTH_SOLIDUS Or c = FULLWIDTH_HYPHEN_MINUS
        End Function


        Friend Shared Function IsLetterC(ch As Char) As Boolean
            Return _
                ch = "c"c Or ch = "C"c Or ch = FULLWIDTH_LATIN_CAPITAL_LETTER_C Or ch = FULLWIDTH_LATIN_SMALL_LETTER_C
        End Function

        ''' <summary>
        ''' matches one char or another.
        ''' Typical usage is for matching lowercase and uppercase.
        ''' </summary>
        Friend Shared Function MatchOneOrAnother(ch As Char, one As Char, another As Char) As Boolean
            Return ch = one Or ch = another
        End Function

        ''' <summary>
        ''' matches one char or another.
        ''' it will try normal width and then fullwidth variations.
        ''' Typical usage is for matching lowercase and uppercase.
        ''' </summary>
        Friend Shared Function MatchOneOrAnotherOrFullwidth(
                                                             ch As Char,
                                                             one As Char,
                                                             another As Char
                                                           ) As Boolean
            Debug.Assert(IsHalfWidth(one))
            Debug.Assert(IsHalfWidth(another))

            If IsFullWidth(ch) Then ch = MakeHalfWidth(ch)
            Return ch = one Or ch = another
        End Function

        Friend Shared Function IsPropAlpha(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties <= UnicodeCategory.OtherLetter
        End Function

        Friend Shared Function IsPropAlphaNumeric(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties <= UnicodeCategory.DecimalDigitNumber
        End Function

        Friend Shared Function IsPropLetterDigit(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties = UnicodeCategory.LetterNumber
        End Function

        Friend Shared Function IsPropConnectorPunctuation(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties = UnicodeCategory.ConnectorPunctuation
        End Function

        Friend Shared Function IsPropCombining(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties >= UnicodeCategory.NonSpacingMark AndAlso
                CharacterProperties <= UnicodeCategory.EnclosingMark
        End Function

        Friend Shared Function IsConnectorPunctuation(c As Char) As Boolean
            Return CharUnicodeInfo.GetUnicodeCategory(c) = UnicodeCategory.ConnectorPunctuation
        End Function

        Friend Shared Function IsSpaceSeparator(c As Char) As Boolean
            Return CharUnicodeInfo.GetUnicodeCategory(c) = UnicodeCategory.SpaceSeparator
        End Function

        Friend Shared Function IsPropOtherFormat(CharacterProperties As UnicodeCategory) As Boolean
            Return CharacterProperties = UnicodeCategory.Format
        End Function

        Friend Shared Function IsSurrogate(c As Char) As Boolean
            Return Char.IsSurrogate(c)
        End Function

        Friend Shared Function IsHighSurrogate(c As Char) As Boolean
            Return Char.IsHighSurrogate(c)
        End Function

        Friend Shared Function IsLowSurrogate(c As Char) As Boolean
            Return Char.IsLowSurrogate(c)
        End Function

        Friend Shared Function ReturnFullWidthOrSelf(c As Char) As Char
            Return If( IsHalfWidth(c), MakeFullWidth(c), c )
        End Function
    End Class

End Namespace

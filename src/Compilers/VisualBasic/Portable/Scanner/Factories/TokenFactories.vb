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

#Region "Caches"

#Region "Trivia"
        Private Structure TriviaKey
            Public ReadOnly spelling As String
            Public ReadOnly kind As SyntaxKind

            Public Sub New(spelling As String, kind As SyntaxKind)
                Me.spelling = spelling
                Me.kind = kind
            End Sub
        End Structure

        Private Shared ReadOnly s_triviaKeyHasher As Func(Of TriviaKey, Integer) =
            Function(key) RuntimeHelpers.GetHashCode(key.spelling) Xor key.kind

        Private Shared ReadOnly s_triviaKeyEquality As Func(Of TriviaKey, SyntaxTrivia, Boolean) =
            Function(key, value) (key.spelling Is value.Text) AndAlso (key.kind = value.Kind)

        Private Shared ReadOnly s_singleSpaceWhitespaceTrivia   As SyntaxTrivia = SyntaxFactory.WhitespaceTrivia(" ")
        Private Shared ReadOnly s_fourSpacesWhitespaceTrivia    As SyntaxTrivia = SyntaxFactory.WhitespaceTrivia("    ")
        Private Shared ReadOnly s_eightSpacesWhitespaceTrivia   As SyntaxTrivia = SyntaxFactory.WhitespaceTrivia("        ")
        Private Shared ReadOnly s_twelveSpacesWhitespaceTrivia  As SyntaxTrivia = SyntaxFactory.WhitespaceTrivia("            ")
        Private Shared ReadOnly s_sixteenSpacesWhitespaceTrivia As SyntaxTrivia = SyntaxFactory.WhitespaceTrivia("                ")


        Private Shared Function CreateWsTable() As CachingFactory(Of TriviaKey, SyntaxTrivia)
            Dim table As New CachingFactory(Of TriviaKey, SyntaxTrivia)(TABLE_LIMIT, Nothing, s_triviaKeyHasher, s_triviaKeyEquality)

            ' Prepopulate the table with some common values
            table.Add(New TriviaKey(" ", SyntaxKind.WhitespaceTrivia), s_singleSpaceWhitespaceTrivia)
            table.Add(New TriviaKey("    ", SyntaxKind.WhitespaceTrivia), s_fourSpacesWhitespaceTrivia)
            table.Add(New TriviaKey("        ", SyntaxKind.WhitespaceTrivia), s_eightSpacesWhitespaceTrivia)
            table.Add(New TriviaKey("            ", SyntaxKind.WhitespaceTrivia), s_twelveSpacesWhitespaceTrivia)
            table.Add(New TriviaKey("                ", SyntaxKind.WhitespaceTrivia), s_sixteenSpacesWhitespaceTrivia)

            Return table
        End Function

#End Region

#Region "WhiteSpaceList"

        Private Shared ReadOnly s_wsListKeyHasher As Func(Of SyntaxListBuilder, Integer) =
            Function(builder)
                Dim code = 0
                For i = 0 To builder.Count - 1
                    Dim value = builder(i)
                    ' shift because there could be the same trivia nodes in the list
                    code = (code << 1) Xor RuntimeHelpers.GetHashCode(value)
                Next
                Return code
            End Function

        Private Shared ReadOnly s_wsListKeyEquality As Func(Of SyntaxListBuilder, CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), Boolean) =
            Function(builder, list)
                If builder.Count <> list.Count Then Return False
                For i = 0 To builder.Count - 1
                    If builder(i) IsNot list.ItemUntyped(i) Then Return False
                Next
                Return True
            End Function

        Private Shared ReadOnly s_wsListFactory As Func(Of SyntaxListBuilder, CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) =
            Function(builder) builder.ToList(Of VisualBasicSyntaxNode)()


#End Region

#Region "Tokens"

        Private Structure TokenParts
            Friend ReadOnly spelling As String
            Friend ReadOnly pTrivia As GreenNode
            Friend ReadOnly fTrivia As GreenNode

            Friend Sub New(pTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), fTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), spelling As String)
                Me.spelling = spelling
                Me.pTrivia = pTrivia.Node
                Me.fTrivia = fTrivia.Node
            End Sub
        End Structure

        Private Shared ReadOnly s_tokenKeyHasher As Func(Of TokenParts, Integer) =
            Function(key)
                Dim code = RuntimeHelpers.GetHashCode(key.spelling)
                Dim trivia = key.pTrivia
                If trivia IsNot Nothing Then  code = code Xor (RuntimeHelpers.GetHashCode(trivia) << 1)
                trivia = key.fTrivia
                If trivia IsNot Nothing Then  code = code Xor RuntimeHelpers.GetHashCode(trivia)
                Return code
            End Function

        Private Shared ReadOnly s_tokenKeyEquality As Func(Of TokenParts, SyntaxToken, Boolean) =
            Function(x, y)
                If y Is Nothing OrElse
                   x.spelling IsNot y.Text OrElse
                   x.fTrivia IsNot y.GetTrailingTrivia OrElse
                   x.pTrivia IsNot y.GetLeadingTrivia Then

                    Return False
                End If

                Return True
            End Function

#End Region

        Private Shared Function CanCache(trivia As SyntaxListBuilder) As Boolean
            For i = 0 To trivia.Count - 1
                Dim t = trivia(i)
                Select Case t.RawKind
                    Case SyntaxKind.WhitespaceTrivia,
                         SyntaxKind.EndOfLineTrivia,
                         SyntaxKind.LineContinuationTrivia,
                         SyntaxKind.DocumentationCommentExteriorTrivia

                        'do nothing
                    Case Else
                        Return False
                End Select
            Next
            Return True
        End Function

#End Region

#Region "Trivia"

        Private Function MakeTrivia(
                                     text As String,
                                     kind As SyntaxKind,
                                     f As Func(Of String, SyntaxTrivia) ) As SyntaxTrivia
            Dim ws As SyntaxTrivia = Nothing
            Dim key as New TriviaKey(text, kind)
            If Not _wsTable.TryGetValue(key, ws) Then
                ws = f(text)
                _wsTable.Add(key, ws)
            End If
            Return ws
        End Function

        Friend Function MakeWhiteSpaceTrivia(text As String) As SyntaxTrivia
            Debug.Assert(text.Length > 0)
            Debug.Assert(text.All(AddressOf IsWhitespace))
            Return MakeTrivia(text,SyntaxKind.WhitespaceTrivia, AddressOf SyntaxFactory.WhitespaceTrivia)
        End Function

        Friend Function MakeEndOfLineTrivia(text As String) As SyntaxTrivia
            Return MakeTrivia(text,SyntaxKind.EndOfLineTrivia, AddressOf SyntaxFactory.EndOfLineTrivia)
        End Function

        Friend Function MakeColonTrivia(text As String) As SyntaxTrivia
            Debug.Assert(text.Length = 1)
            Debug.Assert(IsColon(text(0)))
            Return MakeTrivia(text, SyntaxKind.ColonTrivia, AddressOf SyntaxFactory.ColonTrivia)
        End Function

        Private Shared ReadOnly s_crLfTrivia As SyntaxTrivia = SyntaxFactory.EndOfLineTrivia(vbCrLf)
        Friend Function MakeEndOfLineTriviaCRLF() As SyntaxTrivia
            AdvanceChar(2)
            Return s_crLfTrivia
        End Function

        Friend Function MakeLineContinuationTrivia(text As String) As SyntaxTrivia
            Debug.Assert(text.Length = 1)
            Debug.Assert(IsUnderscore(text(0)))
            Return MakeTrivia(text,SyntaxKind.LineContinuationTrivia, AddressOf SyntaxFactory.LineContinuationTrivia)
        End Function

        Friend Function MakeDocumentationCommentExteriorTrivia(text As String) As SyntaxTrivia
            Return MakeTrivia(text, SyntaxKind.DocumentationCommentExteriorTrivia, AddressOf SyntaxFactory.DocumentationCommentExteriorTrivia)
        End Function

        Friend Shared Function MakeCommentTrivia(text As String) As SyntaxTrivia
            Return SyntaxFactory.SyntaxTrivia(SyntaxKind.CommentTrivia, text)
        End Function

        Friend Function MakeTriviaArray(builder As SyntaxListBuilder) As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
            If builder.Count = 0 Then Return Nothing
            Dim foundTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) = Nothing
            Dim useCache = CanCache(builder)
            If useCache Then  Return _wslTable.GetOrMakeValue(builder)
            Return builder.ToList
        End Function

#End Region

#Region "Identifiers"
        Private Function MakeIdentifier(spelling As String,
                                       contextualKind As SyntaxKind,
                                       isBracketed As Boolean,
                                       BaseSpelling As String,
                                       TypeCharacter As TypeCharacter,
                                       leadingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As IdentifierTokenSyntax

            Dim followingTrivia = ScanSingleLineTrivia()

            Return MakeIdentifier(spelling,
                               contextualKind,
                               isBracketed,
                               BaseSpelling,
                               TypeCharacter,
                               leadingTrivia,
                               followingTrivia)

        End Function

        Friend Function MakeIdentifier(keyword As KeywordSyntax) As IdentifierTokenSyntax
            Return MakeIdentifier(keyword.Text,
                                  keyword.Kind,
                                  False,
                                  keyword.Text,
                                  TypeCharacter.None,
                                  keyword.GetLeadingTrivia,
                                  keyword.GetTrailingTrivia)
        End Function

        Private Function MakeIdentifier(spelling As String,
                               contextualKind As SyntaxKind,
                               isBracketed As Boolean,
                               BaseSpelling As String,
                               TypeCharacter As TypeCharacter,
                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode),
                               followingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode)) As IdentifierTokenSyntax

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim id As IdentifierTokenSyntax = Nothing
            If _idTable.TryGetValue(tp, id) Then Return id

            If contextualKind <> SyntaxKind.IdentifierToken OrElse
                isBracketed = True OrElse
                TypeCharacter <> TypeCharacter.None Then

                id = SyntaxFactory.Identifier(spelling, contextualKind, isBracketed, BaseSpelling, TypeCharacter, precedingTrivia.Node, followingTrivia.Node)
            Else
                id = SyntaxFactory.Identifier(spelling, precedingTrivia.Node, followingTrivia.Node)
            End If

            _idTable.Add(tp, id)

            Return id
        End Function
#End Region

#Region "Keywords"

        Private Function MakeKeyword(
                                      tokenType As SyntaxKind,
                                      spelling As String,
                                      precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                    ) As KeywordSyntax

            Dim followingTrivia = ScanSingleLineTrivia()

            Return MakeKeyword(tokenType,
                               spelling,
                               precedingTrivia,
                               followingTrivia)
        End Function

        Friend Function MakeKeyword(identifier As IdentifierTokenSyntax) As KeywordSyntax
            Debug.Assert(identifier.PossibleKeywordKind <> SyntaxKind.IdentifierToken AndAlso
                         Not identifier.IsBracketed AndAlso
                         (identifier.TypeCharacter = TypeCharacter.None OrElse identifier.PossibleKeywordKind = SyntaxKind.MidKeyword))

            Return MakeKeyword(identifier.PossibleKeywordKind,
                               identifier.Text,
                               identifier.GetLeadingTrivia,
                               identifier.GetTrailingTrivia)
        End Function

        Friend Function MakeKeyword(xmlName As XmlNameTokenSyntax) As KeywordSyntax
            Debug.Assert(xmlName.PossibleKeywordKind <> SyntaxKind.XmlNameToken)

            Return MakeKeyword(xmlName.PossibleKeywordKind,
                               xmlName.Text,
                               xmlName.GetLeadingTrivia,
                               xmlName.GetTrailingTrivia)
        End Function

        Private Function MakeKeyword(
                                      tokenType As SyntaxKind,
                                      spelling As String,
                                      precedingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode),
                                      followingTrivia As CoreInternalSyntax.SyntaxList(Of GreenNode)
                                    ) As KeywordSyntax

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)
            Dim kw As KeywordSyntax = Nothing
            If _kwTable.TryGetValue(tp, kw) Then  Return kw
            kw = New KeywordSyntax(tokenType, spelling, precedingTrivia.Node, followingTrivia.Node)
            _kwTable.Add(tp, kw)
            Return kw
        End Function
#End Region


#Region "Literals"
        Private Function MakeIntegerLiteralToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                         base As LiteralBase,
                                         typeCharacter As TypeCharacter,
                                         integralValue As ULong,
                                         length As Integer) As SyntaxToken

            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim p As SyntaxToken = Nothing
            If _literalTable.TryGetValue(tp, p) Then
                Return p
            End If

            p = SyntaxFactory.IntegerLiteralToken(
                        spelling,
                        base,
                        typeCharacter,
                        integralValue,
                        precedingTrivia.Node,
                        followingTrivia.Node)

            _literalTable.Add(tp, p)
            Return p
        End Function

        Private Function MakeCharacterLiteralToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), value As Char, length As Integer) As SyntaxToken
            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim p As SyntaxToken = Nothing
            If _literalTable.TryGetValue(tp, p) Then
                Return p
            End If

            p = SyntaxFactory.CharacterLiteralToken(spelling, value, precedingTrivia.Node, followingTrivia.Node)
            _literalTable.Add(tp, p)
            Return p
        End Function

        Private Function MakeDateLiteralToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), value As DateTime, length As Integer) As SyntaxToken
            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim p As SyntaxToken = Nothing
            If _literalTable.TryGetValue(tp, p) Then
                Return p
            End If

            p = SyntaxFactory.DateLiteralToken(spelling, value, precedingTrivia.Node, followingTrivia.Node)
            _literalTable.Add(tp, p)
            Return p
        End Function

        Private Function MakeFloatingLiteralToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                         typeCharacter As TypeCharacter,
                                         floatingValue As Double,
                                         length As Integer) As SyntaxToken

            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim p As SyntaxToken = Nothing
            If _literalTable.TryGetValue(tp, p) Then
                Return p
            End If

            p = SyntaxFactory.FloatingLiteralToken(
                        spelling,
                        typeCharacter,
                        floatingValue,
                        precedingTrivia.Node,
                        followingTrivia.Node)

            _literalTable.Add(tp, p)
            Return p
        End Function

        Private Function MakeDecimalLiteralToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                 typeCharacter As TypeCharacter,
                                 decimalValue As Decimal,
                                 length As Integer) As SyntaxToken

            Dim spelling = GetText(length)
            Dim followingTrivia = ScanSingleLineTrivia()

            Dim tp As New TokenParts(precedingTrivia, followingTrivia, spelling)

            Dim p As SyntaxToken = Nothing
            If _literalTable.TryGetValue(tp, p) Then
                Return p
            End If

            p = SyntaxFactory.DecimalLiteralToken(
                        spelling,
                        typeCharacter,
                        decimalValue,
                        precedingTrivia.Node,
                        followingTrivia.Node)

            _literalTable.Add(tp, p)
            Return p
        End Function
#End Region

        ' BAD TOKEN

        Private Function MakeBadToken(
                                       precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                       length As Integer,
                                       errId As ERRID
                                     ) As SyntaxToken
            Dim spelling        = GetTextNotInterned(length)
            Dim followingTrivia = ScanSingleLineTrivia()
            Dim result1         = SyntaxFactory.BadToken(SyntaxSubKind.None, spelling, precedingTrivia.Node, followingTrivia.Node)
            Dim errResult1      = result1.AddError(ErrorFactory.ErrorInfo(errId))
            Return DirectCast(errResult1, SyntaxToken)
        End Function

        Private Shared Function MakeEofToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Return SyntaxFactory.Token(precedingTrivia.Node, SyntaxKind.EndOfFileToken, Nothing, String.Empty)
        End Function

        Private Shared ReadOnly _simpleEof As SyntaxToken = SyntaxFactory.Token(Nothing, SyntaxKind.EndOfFileToken, Nothing, String.Empty)
        Private Function MakeEofToken() As SyntaxToken
            Return _simpleEof
        End Function
    End Class
End Namespace

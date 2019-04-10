' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text
'-----------------------------------------------------------------------------

Option Compare Binary
Option Strict On

Imports System.Collections.Immutable
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports Microsoft.CodeAnalysis.Collections
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Features.CheckFeatureAvailability
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Features.LangaugeFeatureService

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    ''' <summary>
    ''' Creates red tokens for a stream of text
    ''' </summary>
    Friend Class Scanner
        Implements IDisposable

        Private Delegate Function ScanTriviaFunc() As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)

        Private Shared ReadOnly s_scanNoTriviaFunc As ScanTriviaFunc = Function() Nothing
        Private ReadOnly _scanSingleLineTriviaFunc As ScanTriviaFunc = AddressOf ScanSingleLineTrivia

        Protected _lineBufferOffset As Integer ' marks the next character to read from _buffer
        Private _endOfTerminatorTrivia As Integer ' marks how far scanner may have scanned ahead for terminator trivia. This may be greater than _lineBufferOffset

        Friend Const BadTokenCountLimit As Integer = 200
        Private _badTokenCount As Integer ' cumulative count of bad tokens produced

        Private ReadOnly _sbPooled As PooledStringBuilder = PooledStringBuilder.GetInstance
        ''' <summary>
        ''' DO NOT USE DIRECTLY.
        ''' USE GetScratch()
        ''' </summary>
        Private ReadOnly _sb As StringBuilder = _sbPooled.Builder
        Private ReadOnly _triviaListPool As New SyntaxListPool
        Private ReadOnly _options As VisualBasicParseOptions

        Private ReadOnly _stringTable As StringTable = StringTable.GetInstance()
        Private ReadOnly _quickTokenTable As TextKeyedCache(Of SyntaxToken) = TextKeyedCache(Of SyntaxToken).GetInstance

        Public Const TABLE_LIMIT = 512
        Private Shared ReadOnly s_keywordKindFactory As Func(Of String, SyntaxKind) =
            Function(spelling) KeywordTable.TokenOfString(spelling)

        Private Shared ReadOnly s_keywordsObjsPool As ObjectPool(Of CachingIdentityFactory(Of String, SyntaxKind)) = CachingIdentityFactory(Of String, SyntaxKind).CreatePool(TABLE_LIMIT, s_keywordKindFactory)
        Private ReadOnly _KeywordsObjs As CachingIdentityFactory(Of String, SyntaxKind) = s_keywordsObjsPool.Allocate()

        Private Shared ReadOnly s_idTablePool As New ObjectPool(Of CachingFactory(Of TokenParts, IdentifierTokenSyntax))(
            Function() New CachingFactory(Of TokenParts, IdentifierTokenSyntax)(TABLE_LIMIT, Nothing, s_tokenKeyHasher, s_tokenKeyEquality))

        Private ReadOnly _idTable As CachingFactory(Of TokenParts, IdentifierTokenSyntax) = s_idTablePool.Allocate()

        Private Shared ReadOnly s_kwTablePool As New ObjectPool(Of CachingFactory(Of TokenParts, KeywordSyntax))(
            Function() New CachingFactory(Of TokenParts, KeywordSyntax)(TABLE_LIMIT, Nothing, s_tokenKeyHasher, s_tokenKeyEquality))

        Private ReadOnly _kwTable As CachingFactory(Of TokenParts, KeywordSyntax) = s_kwTablePool.Allocate

        Private Shared ReadOnly s_punctTablePool As New ObjectPool(Of CachingFactory(Of TokenParts, PunctuationSyntax))(
            Function() New CachingFactory(Of TokenParts, PunctuationSyntax)(TABLE_LIMIT, Nothing, s_tokenKeyHasher, s_tokenKeyEquality))

        Private ReadOnly _punctTable As CachingFactory(Of TokenParts, PunctuationSyntax) = s_punctTablePool.Allocate()

        Private Shared ReadOnly s_literalTablePool As New ObjectPool(Of CachingFactory(Of TokenParts, SyntaxToken))(
            Function() New CachingFactory(Of TokenParts, SyntaxToken)(TABLE_LIMIT, Nothing, s_tokenKeyHasher, s_tokenKeyEquality))

        Private ReadOnly _literalTable As CachingFactory(Of TokenParts, SyntaxToken) = s_literalTablePool.Allocate

        Private Shared ReadOnly s_wslTablePool As New ObjectPool(Of CachingFactory(Of SyntaxListBuilder, CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)))(
            Function() New CachingFactory(Of SyntaxListBuilder, CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode))(TABLE_LIMIT, s_wsListFactory, s_wsListKeyHasher, s_wsListKeyEquality))

        Private ReadOnly _wslTable As CachingFactory(Of SyntaxListBuilder, CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) = s_wslTablePool.Allocate

        Private Shared ReadOnly s_wsTablePool As New ObjectPool(Of CachingFactory(Of TriviaKey, SyntaxTrivia))(
            Function() CreateWsTable())

        Private ReadOnly _wsTable As CachingFactory(Of TriviaKey, SyntaxTrivia) = s_wsTablePool.Allocate

        Private ReadOnly _isScanningForExpressionCompiler As Boolean

        Private _isDisposed As Boolean

        Private Function GetScratch() As StringBuilder
            ' the normal pattern is that we clean scratch after use.
            ' hitting this assert very likely indicates that you
            ' did not release scratch content or worse trying to use
            ' scratch in two places at a time.
            Debug.Assert(_sb.Length = 0, "trying to use dirty buffer?")
            Return _sb
        End Function

#Region "Public interface"
        Friend Sub New(textToScan As SourceText, options As VisualBasicParseOptions, Optional isScanningForExpressionCompiler As Boolean = False)
            Debug.Assert(textToScan IsNot Nothing)

            _lineBufferOffset = 0
            _buffer = textToScan
            _bufferLen = textToScan.Length
            _curPage = GetPage(0)
            _options = options

            _scannerPreprocessorState = New PreprocessorState(GetPreprocessorConstants(options))
            _isScanningForExpressionCompiler = isScanningForExpressionCompiler
        End Sub

        Friend Sub Dispose() Implements IDisposable.Dispose
            If Not _isDisposed Then
                _isDisposed = True

                _KeywordsObjs.Free()
                _quickTokenTable.Free()
                _stringTable.Free()
                _sbPooled.Free()

                s_idTablePool.Free(_idTable)
                s_kwTablePool.Free(_kwTable)
                s_punctTablePool.Free(_punctTable)
                s_literalTablePool.Free(_literalTable)
                s_wslTablePool.Free(_wslTable)
                s_wsTablePool.Free(_wsTable)

                For Each p As Page In _pages
                    If p IsNot Nothing Then
                        p.Free()
                    End If
                Next

                Array.Clear(_pages, 0, _pages.Length)
            End If
        End Sub
        Friend ReadOnly Property Options As VisualBasicParseOptions
            Get
                Return _options
            End Get
        End Property

        Friend Shared Function GetPreprocessorConstants(options As VisualBasicParseOptions) As ImmutableDictionary(Of String, CConst)
            If options.PreprocessorSymbols.IsDefaultOrEmpty Then
                Return ImmutableDictionary(Of String, CConst).Empty
            End If

            Dim result = ImmutableDictionary.CreateBuilder(Of String, CConst)(IdentifierComparison.Comparer)
            For Each symbol In options.PreprocessorSymbols
                ' The values in options have already been verified
                result(symbol.Key) = CConst.CreateChecked(symbol.Value)
            Next

            Return result.ToImmutable()
        End Function

        Private Function GetNextToken(Optional allowLeadingMultilineTrivia As Boolean = False) As SyntaxToken
            ' Use quick token scanning to see if we can scan a token quickly.
            Dim quickToken = QuickScanToken(allowLeadingMultilineTrivia)

            If quickToken.Succeeded Then
                Dim token = _quickTokenTable.FindItem(quickToken.Chars, quickToken.Start, quickToken.Length, quickToken.HashCode)
                If token IsNot Nothing Then
                    AdvanceChar(quickToken.Length)
                    If quickToken.TerminatorLength <> 0 Then
                        _endOfTerminatorTrivia = _lineBufferOffset
                        _lineBufferOffset -= quickToken.TerminatorLength
                    End If

                    Return token
                End If
            End If

            Dim scannedToken = ScanNextToken(allowLeadingMultilineTrivia)

            ' If we quick-scanned a token, but didn't have a actual token cached for it, cache the token we created
            ' from the regular scanner.
            If quickToken.Succeeded Then
                Debug.Assert(quickToken.Length = scannedToken.FullWidth)

                _quickTokenTable.AddItem(quickToken.Chars, quickToken.Start, quickToken.Length, quickToken.HashCode, scannedToken)
            End If

            Return scannedToken
        End Function

        Private Function ScanNextToken(allowLeadingMultilineTrivia As Boolean) As SyntaxToken
#If DEBUG Then
            Dim oldOffset = _lineBufferOffset
#End If
            Dim leadingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)

            If allowLeadingMultilineTrivia Then
                leadingTrivia = ScanMultilineTrivia()
            Else
                leadingTrivia = ScanLeadingTrivia()

                ' Special case where the remainder of the line is a comment.
                Dim length = PeekStartComment(0)
                If length > 0 Then
                    Return MakeEmptyToken(leadingTrivia)
                End If
            End If

            Dim token = TryScanToken(leadingTrivia)

            If token Is Nothing Then
                token = ScanNextCharAsToken(leadingTrivia)
            End If

            If _lineBufferOffset > _endOfTerminatorTrivia Then
                _endOfTerminatorTrivia = _lineBufferOffset
            End If

#If DEBUG Then
            ' we must always consume as much as returned token's full length or things will go very bad
            Debug.Assert(oldOffset + token.FullWidth = _lineBufferOffset OrElse
                         oldOffset + token.FullWidth = _endOfTerminatorTrivia OrElse
                         token.FullWidth = 0)
#End If
            Return token
        End Function

        Private Function ScanNextCharAsToken(leadingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Dim token As SyntaxToken
            Dim c As Char = Nothing
            If Not TryGet(c) Then
                token = MakeEofToken(leadingTrivia)
            Else
                _badTokenCount += 1

                If _badTokenCount < BadTokenCountLimit Then
                    ' // Don't break up surrogate pairs
                    Dim length = If(IsHighSurrogate(c) AndAlso TryGet(c,1) AndAlso IsLowSurrogate(c), 2, 1)
                    token = MakeBadToken(leadingTrivia, length, ERRID.ERR_IllegalChar)
                Else
                    ' If we get too many characters that we cannot make sense of, absorb the rest of the input.
                    token = MakeBadToken(leadingTrivia, RemainingLength(), ERRID.ERR_IllegalChar)
                End If
            End If

            Return token
        End Function

        ' // SkipToNextConditionalLine advances through the input stream until it finds a (logical)
        ' // line that has a '#' character as its first non-whitespace, non-continuation character.
        ' // SkipToNextConditionalLine ignores explicit line continuation.

        ' TODO: this could be vastly simplified if we could ignore line continuations.
        Public Function SkipToNextConditionalLine() As TextSpan
            ' start at current token
            ResetLineBufferOffset()

            Dim start = _lineBufferOffset

            ' if starting not from line start, skip to the next one.
            Dim prev = PrevToken
            If Not IsAtNewLine() OrElse
                (PrevToken IsNot Nothing AndAlso PrevToken.EndsWithEndOfLineOrColonTrivia) Then

                EatThroughLine()
            End If

            Dim condLineStart = _lineBufferOffset
            Dim c As Char = Nothing
            While (TryGet(c))
                Select Case (c)

                    Case CARRIAGE_RETURN, LINE_FEED
                        EatThroughLineBreak(c)
                        condLineStart = _lineBufferOffset
                        Continue While

                    Case SPACE, CHARACTER_TABULATION
                        Debug.Assert(IsWhitespace(Peek()))
                        EatWhitespace()
                        Continue While

                    Case _
                        "a"c, "b"c, "c"c, "d"c, "e"c, "f"c, "g"c, "h"c, "i"c, "j"c, "k"c, "l"c,
                        "m"c, "n"c, "o"c, "p"c, "q"c, "r"c, "s"c, "t"c, "u"c, "v"c, "w"c, "x"c,
                        "y"c, "z"c, "A"c, "B"c, "C"c, "D"c, "E"c, "F"c, "G"c, "H"c, "I"c, "J"c,
                        "K"c, "L"c, "M"c, "N"c, "O"c, "P"c, "Q"c, "R"c, "S"c, "T"c, "U"c, "V"c,
                        "W"c, "X"c, "Y"c, "Z"c, "'"c, "_"c

                        EatThroughLine()
                        condLineStart = _lineBufferOffset
                        Continue While

                    Case "#"c, FULLWIDTH_NUMBER_SIGN
                        Exit While

                    Case Else
                        If IsWhitespace(c) Then
                            EatWhitespace()
                            Continue While

                        ElseIf IsNewLine(c) Then
                            EatThroughLineBreak(c)
                            condLineStart = _lineBufferOffset
                            Continue While

                        End If

                        EatThroughLine()
                        condLineStart = _lineBufferOffset
                        Continue While
                End Select
            End While

            ' we did not find # or we have hit EoF.
            _lineBufferOffset = condLineStart
            Debug.Assert(_lineBufferOffset >= start AndAlso _lineBufferOffset >= 0)

            ResetTokens()
            Return TextSpan.FromBounds(start, condLineStart)
        End Function

        Private Sub EatThroughLine()
            Dim c As Char = Nothing
            While TryGet(c)
                If IsNewLine(c) Then
                    EatThroughLineBreak(c)
                    Return
                Else
                    AdvanceChar()
                End If
            End While
        End Sub

        ''' <summary>
        ''' Gets a chunk of text as a DisabledCode node.
        ''' </summary>
        ''' <param name="span">The range of text.</param>
        ''' <returns>The DisabledCode node.</returns>
        Friend Function GetDisabledTextAt(span As TextSpan) As SyntaxTrivia
            If span.Start >= 0 AndAlso span.End <= _bufferLen Then
                Return SyntaxFactory.DisabledTextTrivia(GetTextNotInterned(span.Start, span.Length))
            End If

            ' TODO: should this be a Require?
            Throw New ArgumentOutOfRangeException(NameOf(span))
        End Function
#End Region

#Region "Interning"
        Friend Function GetScratchTextInterned(sb As StringBuilder) As String
            Dim str = _stringTable.Add(sb)
            sb.Clear()
            Return str
        End Function

        Friend Shared Function GetScratchText(sb As StringBuilder) As String
            ' PERF: Special case for the very common case of a string containing a single space
            Dim str As String
            If sb.Length = 1 AndAlso sb(0) = " "c Then
                str = " "
            Else
                str = sb.ToString
            End If
            sb.Clear()
            Return str
        End Function

        ' This overload of GetScratchText first examines the contents of the StringBuilder to
        ' see if it matches the given string. If so, then the given string is returned, saving
        ' the allocation.
        Private Shared Function GetScratchText(sb As StringBuilder, text As String) As String
            Dim str As String
            If StringTable.TextEquals(text, sb) Then
                str = text
            Else
                str = sb.ToString
            End If
            sb.Clear()
            Return str
        End Function

        Friend Function Intern(s As String, start As Integer, length As Integer) As String
            Return _stringTable.Add(s, start, length)
        End Function

        Friend Function Intern(s As Char(), start As Integer, length As Integer) As String
            Return _stringTable.Add(s, start, length)
        End Function

        Friend Function Intern(ch As Char) As String
            Return _stringTable.Add(ch)
        End Function
        Friend Function Intern(arr As Char()) As String
            Return _stringTable.Add(arr)
        End Function
#End Region

#Region "Buffer helpers"

        Private Function NextAre(chars As String) As Boolean
            Return NextAre(0, chars)
        End Function

        Private Function NextAre(offset As Integer, chars As String) As Boolean
            Debug.Assert(Not String.IsNullOrEmpty(chars))
            Dim n = chars.Length
            If Not CanGet(offset + n - 1) Then Return False
            For i = 0 To n - 1
                If chars(i) <> Peek(offset + i) Then Return False
            Next
            Return True
        End Function

        Private Function NextIs(offset As Integer, c As Char) As Boolean
            Dim ch As Char = Nothing
            Return TryGet(ch, offset) AndAlso (ch = c)
        End Function

        Private Function CanGet(Optional num As Integer = 0) As Boolean
            Debug.Assert(_lineBufferOffset + num >= 0)
            Debug.Assert(num >= -MaxCharsLookBehind)

            Return _lineBufferOffset + num < _bufferLen
        End Function

        Private Function TryGet(ByRef ch As Char, Optional num As Integer = 0) As Boolean
            Dim ok = CanGet(num)
            ch = If(ok, Peek(num), nothing)
            Return ok
        End Function

        Private Function RemainingLength() As Integer
            Dim result = _bufferLen - _lineBufferOffset
            Debug.Assert(CanGet(result - 1))
            Return result
        End Function

        Private Function GetText(length As Integer) As String
            Debug.Assert(length > 0)
            Debug.Assert(CanGet(length - 1))

            If length = 1 Then
                Return GetNextChar()
            End If

            Dim str = GetText(_lineBufferOffset, length)
            AdvanceChar(length)
            Return str
        End Function

        Private Function GetTextNotInterned(length As Integer) As String
            Debug.Assert(length > 0)
            Debug.Assert(CanGet(length - 1))

            If length = 1 Then
                ' we will still intern single chars. There could not be too many.
                Return GetNextChar()
            End If

            Dim str = GetTextNotInterned(_lineBufferOffset, length)
            AdvanceChar(length)
            Return str
        End Function

        Private Sub AdvanceChar(Optional howFar As Integer = 1)
            Debug.Assert(howFar > 0)
            Debug.Assert(CanGet(howFar - 1))

            _lineBufferOffset += howFar
        End Sub

        Private Function GetNextChar() As String
            Debug.Assert(CanGet)

            Dim ch = GetChar()
            _lineBufferOffset += 1

            Return ch
        End Function

        Private Sub EatThroughLineBreak(StartCharacter As Char)
            AdvanceChar(LengthOfLineBreak(StartCharacter))
        End Sub

        Private Function SkipLineBreak(StartCharacter As Char, index As Integer) As Integer
            Return index + LengthOfLineBreak(StartCharacter, index)
        End Function

        Private Function LengthOfLineBreak(StartCharacter As Char, Optional here As Integer = 0) As Integer
            Debug.Assert(CanGet(here))
            Debug.Assert(IsNewLine(StartCharacter))

            Debug.Assert(StartCharacter = Peek(here))

            If StartCharacter = CARRIAGE_RETURN AndAlso NextIs(here + 1, LINE_FEED) Then
                Return 2
            End If
            Return 1
        End Function
#End Region

#Region "New line and explicit line continuation."
        ''' <summary>
        ''' Accept a CR/LF pair or either in isolation as a newline.
        ''' Make it a statement separator
        ''' </summary>
        Private Function ScanNewlineAsStatementTerminator(startCharacter As Char, precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            If _lineBufferOffset >= _endOfTerminatorTrivia Then Return MakeEmptyToken(precedingTrivia)
            Dim width = LengthOfLineBreak(startCharacter)
            Return MakeStatementTerminatorToken(precedingTrivia, width)
        End Function

        Private Function ScanColonAsStatementTerminator(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As SyntaxToken
            If _lineBufferOffset >= _endOfTerminatorTrivia Then Return MakeEmptyToken(precedingTrivia)
            Return MakeColonToken(precedingTrivia, charIsFullWidth)
        End Function

        ''' <summary>
        ''' Accept a CR/LF pair or either in isolation as a newline.
        ''' Make it a whitespace
        ''' </summary>
        Private Function ScanNewlineAsTrivia(StartCharacter As Char) As SyntaxTrivia
            If LengthOfLineBreak(StartCharacter) = 2 Then
                Return MakeEndOfLineTriviaCRLF()
            End If
            Return MakeEndOfLineTrivia(GetNextChar)
        End Function

 
         Private Function IsPreLineContinuationCommentErrorState(ByRef atNewLine As Boolean, ByRef tList As SyntaxListBuilder, ch As Char, Optional Here As Integer = 1) As Boolean
            atNewLine = IsNewLine(ch)
            '   Return Not (atNewLine OrElse Not CanGet(Here))
            If Not atNewLine AndAlso CanGet(Here) Then
                ' If we get here we have an error, return trivia is Nothing
                Return True
            End If
            Return False
        End Function

        Private Sub AddLineContinuationAndOptionalWhitespaces(tList As SyntaxListBuilder, Here As Integer)
            ' Add the Line Continuation Character '_' triva.
            tList.Add(MakeLineContinuationTrivia(GetText(1)))
            If Here <= 1 Then Exit Sub
            ' Add any whitespace trivia.
            tList.Add(MakeWhiteSpaceTrivia(GetText(Here - 1)))
        End Sub

        Private Sub PeekWhitespace(ByRef Here As Integer, ByRef Ch As Char)
            While TryGet(Ch, Here) AndAlso IsWhitespace(Ch)
                Here += 1
            End While
        End Sub

        ''' <summary>
        ''' Scan a line continuation (_) followed by an optional comment
        ''' </summary>
        ''' <param name="tList">A list of trivia starting with _ if returning True
        ''' or Nothing if returning False</param>
        ''' <returns>
        ''' False on errors including
        ''' EOF, Whitespace is missing before _, first character is not _
        ''' True on space _ or _ followed by 0 or more spaces ' Comment even if feature not supported
        ''' If feature is not support error is added to trivia
        ''' </returns>
        Private Function ScanLineContinuation(tList As SyntaxListBuilder) As Boolean
            Dim Here = 1
            Dim atNewLine As Boolean
            Dim ch As Char
            Dim HasOptionalWhiteSpaceOrComment = True
            Dim HasComment = False
            ' Line continuation is valid at the end of the line, or at the end of the file, or followed by a trailing comment.
            ' Eg.  LineContinuation ( EndOfLine | EndOfFile | LineContinuationComment) 
            If Not TryGet(ch) OrElse Not IsAfterWhitespace() OrElse Not IsUnderscore(ch) Then
                Return False
            ElseIf Not TryGet(ch, Here) OrElse (Not IsWhitespace(ch) AndAlso PeekStartComment(Here, AllowREM:=False) <= 0) Then
                ' We don't have a space or ' but we might have an EOF after _ and that is not an error
                If IsPreLineContinuationCommentErrorState(atNewLine, tList, ch) Then
                    Return False
                End If
                HasOptionalWhiteSpaceOrComment = False
            End If
            If HasOptionalWhiteSpaceOrComment Then
                ' We have a Line Continuation
                PeekWhitespace(Here, ch)
                ' followed optional whitespace(s)
                HasComment = PeekStartComment(Here, AllowREM:=False) > 0
                If Not HasComment AndAlso IsPreLineContinuationCommentErrorState(atNewLine, tList, ch, Here) Then
                    '... without a comment.
                    ' so process as pre this feature
                    Return False
                End If
            End If
            AddLineContinuationAndOptionalWhitespaces(tList, Here)
            If HasComment Then
                ' ... with a comment.
                ' Scan the comment trivia.
                Dim comment As SyntaxTrivia = ScanComment(AllowREM:=False)
                comment = CheckFeatureAvailability(comment, Feature.CommentsAfterLineContinuation, Options)
                ' Add the comment trivia.
                tList.Add(comment)
                ch = Peek()
                atNewLine = IsNewLine(ch)
            End If

            ' if there is another line.
            If atNewLine AndAlso CanGet() Then
                ' We need to check it.
                ProcessNextLineToCheckIfBlamkOrBlamkWithComment(tList, ch, Here)
            End If

            Return True
        End Function

        Private Sub ProcessNextLineToCheckIfBlamkOrBlamkWithComment(tList As SyntaxListBuilder, ch As Char, Here As Integer)
            Dim newLine = SkipLineBreak(ch, 0)
            Here = GetWhitespaceLength(newLine)
            Dim spaces = Here - newLine
            Dim startComment = PeekStartComment(Here, AllowREM:=False)
            ' To see if the following the line continuation is ..
            '    * blank
            '    * blank with a comment
            '    So as not confuse code handling Implicit Line Continuations (See Scanner::EatLineContinuation.)
            '    do not include the newline character.
            ' Otherwise include the new line and any additional spaces as trivia.
            If startComment = 0 AndAlso TryGet(ch, Here) AndAlso Not IsNewLine(ch) Then

                tList.Add(MakeEndOfLineTrivia(GetText(newLine)))
                If spaces > 0 Then
                    tList.Add(MakeWhiteSpaceTrivia(GetText(spaces)))
                End If
            End If
        End Sub

#End Region

#Region "Trivia"

        ''' <summary>
        ''' Consumes all trivia until a nontrivia char is found
        ''' </summary>
        Friend Function ScanMultilineTrivia() As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
            Dim ch As Char = Nothing
            If Not TryGet(ch) Then Return Nothing

            ' optimization for a common case
            ' the ASCII range between ': and ~ , with exception of except "'", "_" and R cannot start trivia
            If ch > ":"c AndAlso ch <= "~"c AndAlso Not ch.IsEither("'"c, "_"c, "R"c, "r"c, "<"c, "="c, ">"c) Then
                Return Nothing
            End If

            Dim triviaList = _triviaListPool.Allocate()
            While TryScanSinglePieceOfMultilineTrivia(triviaList)
            End While

            Dim result = MakeTriviaArray(triviaList)
            _triviaListPool.Free(triviaList)
            Return result
        End Function

        ''' <summary>
        ''' Scans a single piece of trivia
        ''' </summary>
        Private Function TryScanSinglePieceOfMultilineTrivia(tList As SyntaxListBuilder) As Boolean
            Dim ch As Char = Nothing
            If Not TryGet(ch) Then Return False

            Dim atNewLine = IsAtNewLine()

            ' check for XmlDocComment and directives
            If atNewLine Then
                If StartsXmlDoc(0)    Then Return TryScanXmlDocComment(tList)
                If StartsDirective(0) Then Return TryScanDirective(tList)
 
                If IsConflictMarkerTrivia() Then
                    ScanConflictMarker(tList)
                    Return True
                End If
            End If

            If IsWhitespace(ch) Then
                ' eat until linebreak or non-whitespace
                Dim wslen = GetWhitespaceLength(1)

                If atNewLine Then
                    If StartsXmlDoc(wslen)    Then Return TryScanXmlDocComment(tList)
                    If StartsDirective(wslen) Then Return TryScanDirective(tList)
                End If
                tList.Add(MakeWhiteSpaceTrivia(GetText(wslen)))
                Return True
            ElseIf IsNewLine(ch) Then
                tList.Add(ScanNewlineAsTrivia(ch))
                Return True
            ElseIf IsUnderscore(ch) Then
                Return ScanLineContinuation(tList)
            ElseIf IsColonAndNotColonEquals(ch, offset:=0) Then
                tList.Add(ScanColonAsTrivia())
                Return True
            End If

            ' try get a comment
            Return ScanCommentIfAny(tList)
        End Function

        ' All conflict markers consist of the same character repeated seven times.  If it Is
        ' a <<<<<<< Or >>>>>>> marker then it Is also followed by a space.
        Private Shared ReadOnly s_conflictMarkerLength As Integer = "<<<<<<<".Length

        Private Function IsConflictMarkerTrivia() As Boolean
            Dim ch as Char = Nothing
            If Not IsAtNewLine OrElse Not TryGet(ch) OrElse Not ch.IsEither("<"c, ">"c, "="c) Then Return False
            If Not (NextAre("<<<<<<<") OrElse NextAre(">>>>>>>") OrElse NextAre("=======")) Then Return False
            If ch = "="c Then Return True
            Return NextIs(s_conflictMarkerLength, " "c)
        End Function

        Private Sub ScanConflictMarker(tList As SyntaxListBuilder)
            Dim startCh = Peek()

            ' First create a trivia from the start of this merge conflict marker to the
            ' end of line/file (whichever comes first).
            ScanConflictMarkerHeader(tList)

            ' Now add the newlines as the next trivia.
            ScanConflictMarkerEndOfLine(tList)

            If startCh = "="c Then
                ' Consume everything from the start of the mid-conflict marker to the start of the next
                ' end-conflict marker.
                ScanConflictMarkerDisabledText(tList)
            End If
        End Sub

        Private Sub ScanConflictMarkerDisabledText(tList As SyntaxListBuilder)
            Dim start = _lineBufferOffset
            Dim ch As Char = Nothing
            While TryGet(ch) AndAlso Not (ch = ">"c AndAlso IsConflictMarkerTrivia())
                AdvanceChar()
            End While

            Dim width = _lineBufferOffset - start
            If width > 0 Then tList.Add(SyntaxFactory.DisabledTextTrivia(GetText(start, width)))
        End Sub

        Private Sub ScanConflictMarkerEndOfLine(tList As SyntaxListBuilder)
            Dim start = _lineBufferOffset
            Dim ch AS Char = Nothing
            While TryGet(ch) AndAlso SyntaxFacts.IsNewLine(ch)
                AdvanceChar()
            End While

            Dim width = _lineBufferOffset - start
            If width > 0 Then tList.Add(SyntaxFactory.EndOfLineTrivia(GetText(start, width)))
        End Sub

        Private Sub ScanConflictMarkerHeader(tList As SyntaxListBuilder)
            Dim start = _lineBufferOffset
            DIm ch As Char = Nothing
            While TryGet(ch) AndAlso Not SyntaxFacts.IsNewLine(ch)
                AdvanceChar()
            End While

            Dim trivia = SyntaxFactory.ConflictMarkerTrivia(GetText(start, _lineBufferOffset - start))
            trivia = DirectCast(trivia.SetDiagnostics({ErrorFactory.ErrorInfo(ERRID.ERR_Merge_conflict_marker_encountered)}), SyntaxTrivia)
            tList.Add(trivia)
        End Sub

        ' check for '''(~')
        Private Function StartsXmlDoc(Here As Integer) As Boolean
            Return _options.DocumentationMode >= DocumentationMode.Parse AndAlso
                CanGet(Here + 3) AndAlso
                IsSingleQuote(Peek(Here)) AndAlso
                IsSingleQuote(Peek(Here + 1)) AndAlso
                IsSingleQuote(Peek(Here + 2)) AndAlso
                Not IsSingleQuote(Peek(Here + 3))
        End Function

        ' check for #
        Private Function StartsDirective(Here As Integer) As Boolean
            Dim ch As Char = Nothing
            Return If(TryGet(ch, Here), IsHash(ch), False)
        End Function

        Private Function IsAtNewLine() As Boolean
            Return _lineBufferOffset = 0 OrElse IsNewLine(PrevChar())
        End Function

        Private Function IsAfterWhitespace() As Boolean
            Return (_lineBufferOffset = 0) OrElse IsWhitespace(PrevChar())
        End Function

        ''' <summary>
        ''' Scan trivia on one LOGICAL line
        ''' Will check for whitespace, comment, EoL, implicit line break
        ''' EoL may be consumed as whitespace only as a part of line continuation ( _ )
        ''' </summary>
        Friend Function ScanSingleLineTrivia() As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
            Dim tList = _triviaListPool.Allocate()
            ScanSingleLineTrivia(tList)
            Dim result = MakeTriviaArray(tList)
            _triviaListPool.Free(tList)
            Return result
        End Function

        Private Sub ScanSingleLineTrivia(tList As SyntaxListBuilder)
            If IsScanningXmlDoc Then
                ScanSingleLineTriviaInXmlDoc(tList)
            Else
                ScanWhitespaceAndLineContinuations(tList)
                ScanCommentIfAny(tList)
                ScanTerminatorTrivia(tList)
            End If
        End Sub

        Private Sub ScanSingleLineTriviaInXmlDoc(tList As SyntaxListBuilder)
            Dim c As Char = Nothing
            If Not TryGet(c) Then Exit Sub
            Select Case (c)
                ' // Whitespace
                ' //  S    ::=    (#x20 | #x9 | #xD | #xA)+
                Case CARRIAGE_RETURN, LINE_FEED, " "c, CHARACTER_TABULATION
                    Dim offsets = CreateOffsetRestorePoint()
                    Dim triviaList = _triviaListPool.Allocate(Of VisualBasicSyntaxNode)()
                    Dim continueLine = ScanXmlTriviaInXmlDoc(c, triviaList)
                    If Not continueLine Then
                        _triviaListPool.Free(triviaList)
                        offsets.Restore()
                        Return
                    End If

                    tlist.AddRange(MakeTriviaArray(triviaList))

                    _triviaListPool.Free(triviaList)

            End Select
        End Sub

        Private Function ScanLeadingTrivia() As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
            Dim tList = _triviaListPool.Allocate()
            ScanWhitespaceAndLineContinuations(tList)
            Dim result = MakeTriviaArray(tList)
            _triviaListPool.Free(tList)
            Return result
        End Function

        Private Sub ScanWhitespaceAndLineContinuations(tList As SyntaxListBuilder)
            Dim ch As Char = Nothing
            If TryGet(ch) AndAlso IsWhitespace(ch) Then
                tList.Add(ScanWhitespace(1))
                ' collect { lineCont, ws }
                While ScanLineContinuation(tList)
                End While
            End If
        End Sub

        Private Function ScanSingleLineTrivia(includeFollowingBlankLines As Boolean) As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
            Dim tList = _triviaListPool.Allocate()
            ScanSingleLineTrivia(tList)

            If includeFollowingBlankLines AndAlso IsBlankLine(tList) Then
                Dim more = _triviaListPool.Allocate()

                While True
                    Dim offsets = CreateOffsetRestorePoint()

                    _lineBufferOffset = _endOfTerminatorTrivia
                    ScanSingleLineTrivia(more)

                    If Not IsBlankLine(more) Then
                        offsets.Restore()
                        Exit While
                    End If

                    tlist.AddRange(MakeTriviaArray(more))

                    more.Clear()
                End While

                _triviaListPool.Free(more)
            End If

            Dim result = tList.ToList()
            _triviaListPool.Free(tList)
            Return result
        End Function

        ''' <summary>
        ''' Return True if the builder is a (possibly empty) list of
        ''' WhitespaceTrivia followed by an EndOfLineTrivia.
        ''' </summary>
        Private Shared Function IsBlankLine(tList As SyntaxListBuilder) As Boolean
            Dim n = tList.Count
            If n = 0 OrElse tList(n - 1).RawKind <> SyntaxKind.EndOfLineTrivia Then
                Return False
            End If
            For i = 0 To n - 2
                If tList(i).RawKind <> SyntaxKind.WhitespaceTrivia Then
                    Return False
                End If
            Next
            Return True
        End Function

        Private Sub ScanTerminatorTrivia(tList As SyntaxListBuilder)
            ' Check for statement terminators
            ' There are 4 special cases

            '   1. [colon ws+]* colon -> colon terminator
            '   2. new line -> new line terminator
            '   3. colon followed by new line -> colon terminator + new line terminator
            '   4. new line followed by new line -> new line terminator + new line terminator

            ' Case 3 is required to parse single line if's and numeric labels.
            ' Case 4 is required to limit explicit line continuations to single new line
            Dim ch As Char = Nothing
            If Not TryGet(ch) Then Exit Sub
            Dim startOfTerminatorTrivia = _lineBufferOffset

            If IsNewLine(ch) Then
                tList.Add(ScanNewlineAsTrivia(ch))

            ElseIf IsColonAndNotColonEquals(ch, offset:=0) Then
                tList.Add(ScanColonAsTrivia())

                ' collect { ws, colon }
                Do
                    Dim len = GetWhitespaceLength(0)
                    If Not TryGet(ch, len) OrElse Not IsColonAndNotColonEquals(ch, offset:=len) Then Exit Do

                    If len > 0 Then tList.Add(MakeWhiteSpaceTrivia(GetText(len)))
 
                    startOfTerminatorTrivia = _lineBufferOffset
                    tList.Add(ScanColonAsTrivia())
                Loop
            End If

            _endOfTerminatorTrivia = _lineBufferOffset
            ' Reset _lineBufferOffset to the start of the terminator trivia.
            ' When the scanner is asked for the next token, it will return a 0 length terminator or colon token.
            _lineBufferOffset = startOfTerminatorTrivia
        End Sub

        Private Function ScanCommentIfAny(tList As SyntaxListBuilder) As Boolean
            If CanGet() Then
                ' check for comment
                Dim comment = ScanComment()
                If comment IsNot Nothing Then
                    tList.Add(comment)
                    Return True
                End If
            End If
            Return False
        End Function

        Private Function GetWhitespaceLength(len As Integer) As Integer
            Dim ch As Char = Nothing
            ' eat until linebreak or non-whitespace
            While TryGet(ch, len) AndAlso IsWhitespace(ch)
                len += 1
            End While
            Return len
        End Function

        Private Function GetXmlWhitespaceLength(len As Integer) As Integer
            Dim ch As Char = Nothing
            ' eat until linebreak or non-whitespace
            While TryGet(ch, len) AndAlso IsXmlWhitespace(ch)
                len += 1
            End While
            Return len
        End Function

        Private Function ScanWhitespace(Optional len As Integer = 0) As VisualBasicSyntaxNode
            len = GetWhitespaceLength(len)
            Return If( len > 0, MakeWhiteSpaceTrivia(GetText(len)), Nothing)
        End Function

        Private Function ScanXmlWhitespace(Optional len As Integer = 0) As VisualBasicSyntaxNode
            len = GetXmlWhitespaceLength(len)
            Return If(len > 0, MakeWhiteSpaceTrivia(GetText(len)), Nothing)
        End Function

        Private Sub EatWhitespace()
            Debug.Assert(CanGet)
            Debug.Assert(IsWhitespace(Peek()))

            AdvanceChar()

            Dim ch As Char = Nothing
            ' eat until linebreak or non-whitespace
            While TryGet(ch) AndAlso IsWhitespace(ch)
                AdvanceChar()
            End While
        End Sub

       Private Function PeekStartComment(i As Integer, Optional AllowREM As Boolean = True) As Integer
            Dim ch AS Char
            If TryGet(ch, i) Then
                If IsSingleQuote(ch) Then
                    Return 1
                ElseIf AllowREM AndAlso
                    MatchOneOrAnotherOrFullwidth(ch, "R"c, "r"c) AndAlso
                    TryGet(ch, i + 2) AndAlso MatchOneOrAnotherOrFullwidth(Peek(i + 1), "E"c, "e"c) AndAlso
                    MatchOneOrAnotherOrFullwidth(ch, "M"c, "m"c) Then
                    Dim ok = TryGet(ch, i + 3)
                    If Not ok OrElse IsNewLine(ch) Then
                        ' have only 'REM'
                        Return 3
                    ElseIf Not IsIdentifierPartCharacter(ch) Then
                        ' have 'REM '
                        Return 4
                    End If
                End If
            End If

            Return 0
        End Function

       Private Function ScanComment(Optional AllowREM As Boolean = True) As SyntaxTrivia
         Debug.Assert(CanGet())

         Dim length = PeekStartComment(0, AllowREM)
         If length <= 0 Then Return Nothing
         Dim looksLikeDocComment As Boolean = StartsXmlDoc(0)
         Dim ch As Char = Nothing
         ' eat all chars until EoL
         While TryGet(ch, length) AndAlso Not IsNewLine(ch)
           length += 1
         End While

         Dim commentTrivia As SyntaxTrivia = MakeCommentTrivia(GetTextNotInterned(length))
         If looksLikeDocComment AndAlso _options.DocumentationMode >= DocumentationMode.Diagnose Then
            commentTrivia = commentTrivia.WithDiagnostics(ErrorFactory.ErrorInfo(ERRID.WRN_XMLDocNotFirstOnLine))
         End If
         Return commentTrivia
        End Function

        ''' <summary>
        ''' Return True if the character is a colon, and not part of ":=".
        ''' </summary>
        Private Function IsColonAndNotColonEquals(ch As Char, offset As Integer) As Boolean
            Return IsColon(ch) AndAlso Not TrySkipFollowingEquals(offset + 1)
        End Function

        Private Function ScanColonAsTrivia() As SyntaxTrivia
            Debug.Assert(CanGet())
            Debug.Assert(IsColonAndNotColonEquals(Peek(), offset:=0))

            Return MakeColonTrivia(GetText(1))
        End Function

#End Region

        Private Function ScanTokenCommon(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), ch As Char, fullWidth As Boolean) As SyntaxToken
            Dim lengthWithMaybeEquals As Integer = 1
            Select Case ch
                Case CARRIAGE_RETURN, LINE_FEED
                    Return ScanNewlineAsStatementTerminator(ch, precedingTrivia)

                Case NEXT_LINE, LINE_SEPARATOR, PARAGRAPH_SEPARATOR
                    If Not fullWidth Then
                        Return ScanNewlineAsStatementTerminator(ch, precedingTrivia)
                    End If

                Case " "c, CHARACTER_TABULATION, "'"c
                    Debug.Assert(False, $"Unexpected char: &H{AscW(ch):x}")
                    Return Nothing ' trivia cannot start a token

                Case "@"c
                    Return MakeAtToken(precedingTrivia, fullWidth)

                Case "("c
                    Return MakeOpenParenToken(precedingTrivia, fullWidth)

                Case ")"c
                    Return MakeCloseParenToken(precedingTrivia, fullWidth)

                Case "{"c
                    Return MakeOpenBraceToken(precedingTrivia, fullWidth)

                Case "}"c
                    Return MakeCloseBraceToken(precedingTrivia, fullWidth)

                Case ","c
                    Return MakeCommaToken(precedingTrivia, fullWidth)

                Case "#"c
                    Dim dl = ScanDateLiteral(precedingTrivia)
                    Return If(dl, MakeHashToken(precedingTrivia, fullWidth))

                Case "&"c
                    If TryGet(ch, 1) AndAlso BeginsBaseLiteral(ch) Then
                        Return ScanNumericLiteral(precedingTrivia)
                    End If

                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeAmpersandEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeAmpersandToken(precedingTrivia, fullWidth)
                    End If

                Case "="c
                    Return MakeEqualsToken(precedingTrivia, fullWidth)

                Case "<"c
                    Return ScanLeftAngleBracket(precedingTrivia, fullWidth, _scanSingleLineTriviaFunc)

                Case ">"c
                    Return ScanRightAngleBracket(precedingTrivia, fullWidth)

                Case ":"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeColonEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return ScanColonAsStatementTerminator(precedingTrivia, fullWidth)
                    End If

                Case "+"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakePlusEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakePlusToken(precedingTrivia, fullWidth)
                    End If

                Case "-"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeMinusEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeMinusToken(precedingTrivia, fullWidth)
                    End If

                Case "*"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeAsteriskEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeAsteriskToken(precedingTrivia, fullWidth)
                    End If

                Case "/"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeSlashEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeSlashToken(precedingTrivia, fullWidth)
                    End If

                Case "\"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeBackSlashEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeBackslashToken(precedingTrivia, fullWidth)
                    End If

                Case "^"c
                    If TrySkipFollowingEquals(lengthWithMaybeEquals) Then
                        Return MakeCaretEqualsToken(precedingTrivia, lengthWithMaybeEquals)
                    Else
                        Return MakeCaretToken(precedingTrivia, fullWidth)
                    End If

                Case "!"c
                    Dim nc As Char = Nothing
                    ' Check to see if could be a FlagsEnum Operator.
                    If TryGet(nc, 1) Then
                        Select Case nc
                            Case "+"c
                                AdvanceChar(2)
                                Return SyntaxFactory.FlagsEnumSetToken("!+", precedingTrivia.Node, Nothing)
                            Case "-"c
                                AdvanceChar(2)
                                Return SyntaxFactory.FlagsEnumClearToken("!-", precedingTrivia.Node, Nothing)
                            Case "/"c
                                AdvanceChar(2)
                                Return SyntaxFactory.FlagsEnumIsAnyToken("!/", precedingTrivia.Node, Nothing)
                        End Select
                    End If
                    Return MakeExclamationToken(precedingTrivia, fullWidth)

                Case "."c
                    If TryGet(ch, 1) AndAlso IsDecimalDigit(ch) Then
                        Return ScanNumericLiteral(precedingTrivia)
                    Else
                        Return MakeDotToken(precedingTrivia, fullWidth)
                    End If

                Case "0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c
                    Return ScanNumericLiteral(precedingTrivia)

                Case """"c
                    Return ScanStringLiteral(precedingTrivia)

                Case "A"c
                    If NextAre(1, "s ") Then
                        ' TODO: do we allow widechars in keywords?
                        AdvanceChar(2)
                        Return MakeKeyword(SyntaxKind.AsKeyword, "As", precedingTrivia)
                    Else
                        Return ScanIdentifierOrKeyword(precedingTrivia)
                    End If

                Case "E"c
                    If NextAre(1, "nd ") Then
                        ' TODO: do we allow widechars in keywords?
                        AdvanceChar(3)
                        Return MakeKeyword(SyntaxKind.EndKeyword, "End", precedingTrivia)
                    Else
                        Return ScanIdentifierOrKeyword(precedingTrivia)
                    End If

                Case "I"c
                    If NextAre(1, "f ") Then
                        ' TODO: do we allow widechars in keywords?
                        AdvanceChar(2)
                        Return MakeKeyword(SyntaxKind.IfKeyword, "If", precedingTrivia)
                    Else
                        Return ScanIdentifierOrKeyword(precedingTrivia)
                    End If

                Case "a"c, "b"c, "c"c, "d"c, "e"c, "f"c, "g"c, "h"c, "i"c, "j"c, "k"c, "l"c, "m"c,
                     "n"c, "o"c, "p"c, "q"c, "r"c, "s"c, "t"c, "u"c, "v"c, "w"c, "x"c, "y"c, "z"c
                    Return ScanIdentifierOrKeyword(precedingTrivia)

                Case "B"c, "C"c, "D"c, "F"c, "G"c, "H"c, "J"c, "K"c, "L"c, "M"c, "N"c, "O"c, "P"c, "Q"c,
                      "R"c, "S"c, "T"c, "U"c, "V"c, "W"c, "X"c, "Y"c, "Z"c
                    Return ScanIdentifierOrKeyword(precedingTrivia)

                Case "_"c
                    If TryGet(ch, 1) AndAlso IsIdentifierPartCharacter(ch) Then
                        Return ScanIdentifierOrKeyword(precedingTrivia)
                    End If

                    Dim err As ERRID = ERRID.ERR_ExpectedIdentifier
                    Dim len = GetWhitespaceLength(1)
                    If Not TryGet(ch, len) OrElse IsNewLine(ch) OrElse PeekStartComment(len) > 0 Then
                        err = ERRID.ERR_LineContWithCommentOrNoPrecSpace
                    End If

                    ' not a line continuation and cannot start identifier.
                    Return MakeBadToken(precedingTrivia, 1, err)

                Case "["c
                    Return ScanBracketedIdentifier(precedingTrivia)

                Case "?"c
                    Return MakeQuestionToken(precedingTrivia, fullWidth)

                Case "%"c
                    If NextIs(1, ">"c) Then
                        Return XmlMakeEndEmbeddedToken(precedingTrivia, _scanSingleLineTriviaFunc)
                    End If

                Case "$"c, FULLWIDTH_DOLLAR_SIGN
                    Dim nc As Char = Nothing
                    If Not fullWidth AndAlso TryGet(nc, 1) AndAlso IsDoubleQuote(nc) Then
                        Return MakePunctuationToken(precedingTrivia, 2, SyntaxKind.DollarSignDoubleQuoteToken)
                    End If

            End Select
            If IsIdentifierStartCharacter(ch) Then
                Return ScanIdentifierOrKeyword(precedingTrivia)
            End If
            Debug.Assert(Not IsNewLine(ch))
            If fullWidth Then
                Debug.Assert(Not IsDoubleQuote(ch))
                Return Nothing
            End If
            If IsDoubleQuote(ch) Then
                Return ScanStringLiteral(precedingTrivia)
            End If
            If IsFullWidth(ch) Then
                ch = MakeHalfWidth(ch)
                Return ScanTokenFullWidth(precedingTrivia, ch)
            End If
            Return Nothing
        End Function

        ' at this point it is very likely that we are located at the beginning of a token
        Private Function TryScanToken(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            dim ch As Char = Nothing
            Return If(TryGet(ch), ScanTokenCommon(precedingTrivia, ch, False), MakeEofToken(precedingTrivia))
        End Function

        Private Function ScanTokenFullWidth(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), ch As Char) As SyntaxToken
            Return ScanTokenCommon(precedingTrivia, ch, True)
        End Function

        ' // Allow whitespace between the characters of a two-character token.
        Private Function TrySkipFollowingEquals(ByRef Index As Integer) As Boolean
            Debug.Assert(Index > 0)
            Debug.Assert(CanGet(Index - 1))

            Dim Here = Index
            Dim eq As Char

            While TryGet(eq, Here)
                Here += 1
                If Not IsWhitespace(eq) Then
                    If eq.IsEither("="c, FULLWIDTH_EQUALS_SIGN) Then
                        Index = Here
                        Return True
                    Else
                        Return False
                    End If
                End If
            End While
            Return False
        End Function

        Private Function ScanRightAngleBracket(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean) As SyntaxToken
            Debug.Assert(CanGet)  ' >
            Debug.Assert(Peek().IsEither(">"c, FULLWIDTH_GREATER_THAN_SIGN))

            Dim length As Integer = 1

            ' // Allow whitespace between the characters of a two-character token.
            length = GetWhitespaceLength(length)
            Dim c As Char = Nothing
            If TryGet(c, length) Then
                If c.IsEither("="c, FULLWIDTH_EQUALS_SIGN) Then
                    length += 1
                    Return MakeGreaterThanEqualsToken(precedingTrivia, length)
                ElseIf c.IsEither(">"c, FULLWIDTH_GREATER_THAN_SIGN) Then
                    length += 1
                    If TrySkipFollowingEquals(length) Then
                        Return MakeGreaterThanGreaterThanEqualsToken(precedingTrivia, length)
                    Else
                        Return MakeGreaterThanGreaterThanToken(precedingTrivia, length)
                    End If
                End If
            End If
            Return MakeGreaterThanToken(precedingTrivia, charIsFullWidth)
        End Function

        Private Function ScanLeftAngleBracket(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode), charIsFullWidth As Boolean, scanTrailingTrivia As ScanTriviaFunc) As SyntaxToken
            Debug.Assert(CanGet)  ' <
            Debug.Assert(Peek().IsEither("<"c, FULLWIDTH_LESS_THAN_SIGN))

            Dim length As Integer = 1
            Dim c As Char = Nothing
            ' Check for XML tokens
            If Not charIsFullWidth AndAlso TryGet(c, length) Then
                Select Case c
                    Case "!"c
                        If CanGet(length + 2) Then
                            Select Case (Peek(length + 1))
                                Case "-"c
                                    If CanGet(length + 3) AndAlso Peek(length + 2) = "-"c Then
                                        Return XmlMakeBeginCommentToken(precedingTrivia, scanTrailingTrivia)
                                    End If
                                Case "["c

                                    If NextAre(length + 2, "CDATA[") Then

                                        Return XmlMakeBeginCDataToken(precedingTrivia, scanTrailingTrivia)
                                    End If
                            End Select
                        End If
                    Case "?"c
                        Return XmlMakeBeginProcessingInstructionToken(precedingTrivia, scanTrailingTrivia)

                    Case "/"c
                        Return XmlMakeBeginEndElementToken(precedingTrivia, _scanSingleLineTriviaFunc)
                End Select
            End If

            ' // Allow whitespace between the characters of a two-character token.
            length = GetWhitespaceLength(length)

            If TryGet(c, length) Then
                If c.IsEither("="c, FULLWIDTH_EQUALS_SIGN) Then
                    length += 1
                    Return MakeLessThanEqualsToken(precedingTrivia, length)
                ElseIf c.IsEither(">"c, FULLWIDTH_GREATER_THAN_SIGN) Then
                    length += 1
                    Return MakeLessThanGreaterThanToken(precedingTrivia, length)
                ElseIf c.IsEither("<"c, FULLWIDTH_LESS_THAN_SIGN) Then
                    length += 1

                    If TryGet(c, length) Then
                        'if the second "<" is a part of "<%" - like in "<<%" , we do not want to use it.
                        If c <> "%"c AndAlso c <> FULLWIDTH_PERCENT_SIGN Then
                            If TrySkipFollowingEquals(length) Then
                                Return MakeLessThanLessThanEqualsToken(precedingTrivia, length)
                            Else
                                Return MakeLessThanLessThanToken(precedingTrivia, length)
                            End If
                        End If
                    End If
                End If
            End If

            Return MakeLessThanToken(precedingTrivia, charIsFullWidth)
        End Function

        ''' <remarks>
        ''' Not intended for use in Expression Compiler scenarios.
        ''' </remarks>
        Friend Shared Function IsIdentifier(spelling As String) As Boolean
            Dim spellingLength As Integer = spelling.Length
            If spellingLength = 0 Then
                Return False
            End If

            Dim c = spelling(0)
            If SyntaxFacts.IsIdentifierStartCharacter(c) Then
                '  SPEC: ... Visual Basic identifiers conform to the Unicode Standard Annex 15 with one
                '  SPEC:     exception: identifiers may begin with an underscore (connector) character.
                '  SPEC:     If an identifier begins with an underscore, it must contain at least one other
                '  SPEC:     valid identifier character to disambiguate it from a line continuation.
                If IsConnectorPunctuation(c) AndAlso spellingLength = 1 Then
                    Return False
                End If

                For i = 1 To spellingLength - 1
                    If Not IsIdentifierPartCharacter(spelling(i)) Then
                        Return False
                    End If
                Next
            End If

            Return True
        End Function

        Private Function ScanIdentifierOrKeyword(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Debug.Assert(CanGet)
            Debug.Assert(IsIdentifierStartCharacter(Peek))
            Debug.Assert(PeekStartComment(0) = 0) ' comment should be handled by caller

            Dim ch = Peek()
            Dim ch1 As Char = Nothing
            If TryGet(ch1, 1) Then
                If IsConnectorPunctuation(ch) AndAlso Not IsIdentifierPartCharacter(ch1) Then
                    Return MakeBadToken(precedingTrivia, 1, ERRID.ERR_ExpectedIdentifier)
                End If
            End If

            Dim len = 1 ' we know that the first char was good

            ' // The C++ compiler refuses to inline IsIdentifierCharacter, so the
            ' // < 128 test is inline here. (This loop gets a *lot* of traffic.)
            ' TODO: make sure we get good perf here
            While TryGet(ch, len)
                Dim code = Convert.ToUInt16(ch)
                If code < 128US AndAlso IsNarrowIdentifierCharacter(code) OrElse
                    IsWideIdentifierCharacter(ch) Then

                    len += 1
                Else
                    Exit While
                End If
            End While

            'Check for a type character
            Dim TypeCharacter As TypeCharacter = TypeCharacter.None
            If TryGet(ch, len) Then

FullWidthRepeat:
                Select Case ch
                    Case "!"c
                        Dim NextChar As Char = Nothing
                        ' // If the ! is followed by an identifier it is a dictionary lookup operator, not a type character.
                        If TryGet(NextChar, len + 1) Then
                            If IsIdentifierStartCharacter(NextChar) OrElse
                                MatchOneOrAnotherOrFullwidth(NextChar, "["c, "]"c) Then
                                Exit Select
                            ElseIf NextChar.IsEither("+"c, "-"c, "/"c) Then
                                Exit Select
                                'ElseIf NextChar = "("c Then
                                '    Exit Select
                            End If
                        End If
                        TypeCharacter = TypeCharacter.Single  'typeChars.chType_sR4
                        len += 1

                    Case "#"c
                        TypeCharacter = TypeCharacter.Double ' typeChars.chType_sR8
                        len += 1

                    Case "$"c
                        TypeCharacter = TypeCharacter.String 'typeChars.chType_String
                        len += 1

                    Case "%"c
                        TypeCharacter = TypeCharacter.Integer ' typeChars.chType_sI4
                        len += 1

                    Case "&"c
                        TypeCharacter = TypeCharacter.Long 'typeChars.chType_sI8
                        len += 1

                    Case "@"c
                        TypeCharacter = TypeCharacter.Decimal 'chType_sDecimal
                        len += 1

                    Case Else
                        If IsFullWidth(ch) Then
                            ch = MakeHalfWidth(ch)
                            GoTo FullWidthRepeat
                        End If
                End Select
            End If

            Dim tokenType As SyntaxKind = SyntaxKind.IdentifierToken
            Dim contextualKind As SyntaxKind = SyntaxKind.IdentifierToken
            Dim spelling = GetText(len)

            Dim BaseSpelling = If(TypeCharacter = TypeCharacter.None,
                                   spelling,
                                   Intern(spelling, 0, len - 1))

            ' this can be keyword only if it has no type character, or if it is Mid$
            If TypeCharacter = TypeCharacter.None Then
                tokenType = TokenOfStringCached(spelling)
                If SyntaxFacts.IsContextualKeyword(tokenType) Then
                    contextualKind = tokenType
                    tokenType = SyntaxKind.IdentifierToken
                End If
            ElseIf TokenOfStringCached(BaseSpelling) = SyntaxKind.MidKeyword Then

                contextualKind = SyntaxKind.MidKeyword
                tokenType = SyntaxKind.IdentifierToken
            End If

            If tokenType <> SyntaxKind.IdentifierToken Then
                ' KEYWORD
                Return MakeKeyword(tokenType, spelling, precedingTrivia)
            Else
                ' IDENTIFIER or CONTEXTUAL
                Dim id As SyntaxToken = MakeIdentifier(spelling, contextualKind, False, BaseSpelling, TypeCharacter, precedingTrivia)
                Return id
            End If
        End Function

        Private Function TokenOfStringCached(spelling As String) As SyntaxKind
            If spelling.Length > 16 Then
                Return SyntaxKind.IdentifierToken
            End If

            Return _KeywordsObjs.GetOrMakeValue(spelling)
        End Function

        Private Function ScanBracketedIdentifier(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Debug.Assert(CanGet)  ' [
            Debug.Assert(Peek().IsEither("["c, FULLWIDTH_LEFT_SQUARE_BRACKET))

            Dim IdStart As Integer = 1
            Dim Here As Integer = IdStart
            Dim InvalidIdentifier As Boolean = False
            Dim ch AS Char = Nothing
            If Not TryGet(ch, Here) Then Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_MissingEndBrack)

            Dim [Next] As Char = Nothing
            ' check if we can start an ident.
            If Not IsIdentifierStartCharacter(ch) OrElse
                (IsConnectorPunctuation(ch) AndAlso
                    Not (TryGet([Next], Here + 1) AndAlso IsIdentifierPartCharacter([Next]))) Then

                InvalidIdentifier = True
            End If

            ' check ident until ]
            While TryGet([Next], Here)
                If [Next].IsEither("]"c, FULLWIDTH_RIGHT_SQUARE_BRACKET) Then
                    Dim IdStringLength As Integer = Here - IdStart

                    If IdStringLength > 0 AndAlso Not InvalidIdentifier Then
                        Dim spelling = GetText(IdStringLength + 2)
                        ' TODO: this should be provable?
                        Debug.Assert(spelling.Length > IdStringLength + 1)

                        ' TODO: consider interning.
                        Dim baseText = spelling.Substring(1, IdStringLength)
                        Dim id As SyntaxToken = MakeIdentifier(
                            spelling,
                            SyntaxKind.IdentifierToken,
                            True,
                            baseText,
                            TypeCharacter.None,
                            precedingTrivia)
                        Return id
                    Else
                        ' // The sequence "[]" does not define a valid identifier.
                        Return MakeBadToken(precedingTrivia, Here + 1, ERRID.ERR_ExpectedIdentifier)
                    End If
                ElseIf IsNewLine([Next]) Then
                    Exit While
                ElseIf Not IsIdentifierPartCharacter([Next]) Then
                    InvalidIdentifier = True
                    Exit While
                End If

                Here += 1
            End While

            If Here > 1 Then
                Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_MissingEndBrack)
            Else
                Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_ExpectedIdentifier)
            End If
        End Function

        Private Enum NumericLiteralKind
            Integral
            Float
            [Decimal]
        End Enum

        Private Function IsDigitSeperator(ch As Char) As Boolean
            REturn ch = "_"c
        End Function

        Private Function ScanNumericLiteral(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Debug.Assert(CanGet)

            Dim Here As Integer = 0
            Dim IntegerLiteralStart As Integer
            Dim UnderscoreInWrongPlace As Boolean
            Dim UnderscoreUsed As Boolean = False
            Dim LeadingUnderscoreUsed = False

            Dim Base As LiteralBase = LiteralBase.Decimal
            Dim literalKind As NumericLiteralKind = NumericLiteralKind.Integral

            ' ####################################################
            ' // Validate literal and find where the number starts and ends.
            ' ####################################################

            ' // First read a leading base specifier, if present, followed by a sequence of zero
            ' // or more digits.
            Dim ch = Peek()
            If ch.IsEither("&"c, FULLWIDTH_AMPERSAND) Then
                Here += 1
                ch = If(CanGet(Here), Peek(Here), ChrW(0))

FullWidthRepeat:
                Select Case ch
                    Case "H"c, "h"c
                        Here += 1
                        IntegerLiteralStart = Here
                        Base = LiteralBase.Hexadecimal

                        If NextIs(Here, "_"c) Then LeadingUnderscoreUsed = True

                        While TryGet(ch, Here) AndAlso (IsHexDigit(ch) OrElse IsDigitSeperator(ch))
                            If ch = "_"c Then UnderscoreUsed = True
                            Here += 1
                        End While
                        UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)

                    Case "B"c, "b"c
                        Here += 1
                        IntegerLiteralStart = Here
                        Base = LiteralBase.Binary

                        If NextIs(Here, "_"c) Then LeadingUnderscoreUsed = True

                        While TryGet(ch, Here) AndAlso (IsBinaryDigit(ch) OrElse IsDigitSeperator(ch))
                            If ch = "_"c Then UnderscoreUsed = True
                            Here += 1
                        End While
                        UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)

                    Case "O"c, "o"c
                        Here += 1
                        IntegerLiteralStart = Here
                        Base = LiteralBase.Octal

                        If NextIs(Here, "_"c) Then LeadingUnderscoreUsed = True

                        While TryGet(ch, Here) AndAlso (IsOctalDigit(ch) OrElse IsDigitSeperator(ch))
                            If ch = "_"c Then UnderscoreUsed = True
                            Here += 1
                        End While
                        UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)

                    Case Else
                        If IsFullWidth(ch) Then
                            ch = MakeHalfWidth(ch)
                            GoTo FullWidthRepeat
                        End If

                        Throw ExceptionUtilities.UnexpectedValue(ch)
                End Select
            Else
                ' no base specifier - just go through decimal digits.
                IntegerLiteralStart = Here
                UnderscoreInWrongPlace = NextIs(Here, "_"c)
                While TryGet(ch, Here) AndAlso (IsDecimalDigit(ch) OrElse IsDigitSeperator(ch))
                    If ch = "_"c Then UnderscoreUsed = True
                    Here += 1
                End While
                If Here <> IntegerLiteralStart Then
                    UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)
                End If
            End If

            ' we may have a dot, and then it is a float, but if this is an integral, then we have seen it all.
            Dim IntegerLiteralEnd As Integer = Here

            ' // Unless there was an explicit base specifier (which indicates an integer literal),
            ' // read the rest of a float literal.
            If Base = LiteralBase.Decimal AndAlso TryGet(ch, Here) Then
                ' // First read a '.' followed by a sequence of one or more digits.
                Dim nc As Char = Nothing
                If ch.IsEither("."c, FULLWIDTH_FULL_STOP) AndAlso
                    TryGet(nc, Here+1) AndAlso IsDecimalDigit(nc) Then

                    Here += 2   ' skip dot and first digit

                    ' all following decimal digits belong to the literal (fractional part)
                    While TryGet(ch, Here) AndAlso (IsDecimalDigit(ch) OrElse IsDigitSeperator(ch))
                        Here += 1
                    End While
                    UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)
                    literalKind = NumericLiteralKind.Float
                End If

                ' // Read an exponent symbol followed by an optional sign and a sequence of
                ' // one or more digits.
                If TryGet(ch, Here) AndAlso BeginsExponent(nc) Then
                    Here += 1

                    If TryGet(ch, Here) Then
                        If MatchOneOrAnotherOrFullwidth(ch, "+"c, "-"c) Then
                            Here += 1
                        End If
                    End If

                    If TryGet(ch, Here) AndAlso IsDecimalDigit(ch) Then
                        Here += 1
                        While TryGet(ch, Here) AndAlso (IsDecimalDigit(ch) OrElse IsDigitSeperator(ch))
                            Here += 1
                        End While
                        UnderscoreInWrongPlace = UnderscoreInWrongPlace Or NextIs(Here - 1, "_"c)
                    Else
                        Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_InvalidLiteralExponent)
                    End If

                    literalKind = NumericLiteralKind.Float
                End If
            End If

            Dim literalWithoutTypeChar = Here

            ' ####################################################
            ' // Read a trailing type character.
            ' ####################################################

            Dim TypeCharacter As TypeCharacter = TypeCharacter.None

            If TryGet(ch, Here) Then

FullWidthRepeat2:
                Select Case ch
                    Case "!"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.Single
                            literalKind = NumericLiteralKind.Float
                            Here += 1
                        End If

                    Case "F"c, "f"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.SingleLiteral
                            literalKind = NumericLiteralKind.Float
                            Here += 1
                        End If

                    Case "#"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.Double
                            literalKind = NumericLiteralKind.Float
                            Here += 1
                        End If

                    Case "R"c, "r"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.DoubleLiteral
                            literalKind = NumericLiteralKind.Float
                            Here += 1
                        End If

                    Case "S"c, "s"c

                        If literalKind <> NumericLiteralKind.Float Then
                            TypeCharacter = TypeCharacter.ShortLiteral
                            Here += 1
                        End If

                    Case "%"c
                        If literalKind <> NumericLiteralKind.Float Then
                            TypeCharacter = TypeCharacter.Integer
                            Here += 1
                        End If

                    Case "I"c, "i"c
                        If literalKind <> NumericLiteralKind.Float Then
                            TypeCharacter = TypeCharacter.IntegerLiteral
                            Here += 1
                        End If

                    Case "&"c
                        If literalKind <> NumericLiteralKind.Float Then
                            TypeCharacter = TypeCharacter.Long
                            Here += 1
                        End If

                    Case "L"c, "l"c
                        If literalKind <> NumericLiteralKind.Float Then
                            TypeCharacter = TypeCharacter.LongLiteral
                            Here += 1
                        End If

                    Case "@"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.Decimal
                            literalKind = NumericLiteralKind.Decimal
                            Here += 1
                        End If

                    Case "D"c, "d"c
                        If Base = LiteralBase.Decimal Then
                            TypeCharacter = TypeCharacter.DecimalLiteral
                            literalKind = NumericLiteralKind.Decimal

                            ' check if this was not attempt to use obsolete exponent
                            If TryGet(ch, Here + 1) Then
                                If IsDecimalDigit(ch) OrElse MatchOneOrAnotherOrFullwidth(ch, "+"c, "-"c) Then
                                    Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_ObsoleteExponent)
                                End If
                            End If

                            Here += 1
                        End If

                    Case "U"c, "u"c
                        Dim NextChar As Char = Nothing
                        If literalKind <> NumericLiteralKind.Float AndAlso TryGet(NextChar, Here + 1) Then
                            'unsigned suffixes - US, UL, UI
                            If MatchOneOrAnotherOrFullwidth(NextChar, "S"c, "s"c) Then
                                TypeCharacter = TypeCharacter.UShortLiteral
                                Here += 2
                            ElseIf MatchOneOrAnotherOrFullwidth(NextChar, "I"c, "i"c) Then
                                TypeCharacter = TypeCharacter.UIntegerLiteral
                                Here += 2
                            ElseIf MatchOneOrAnotherOrFullwidth(NextChar, "L"c, "l"c) Then
                                TypeCharacter = TypeCharacter.ULongLiteral
                                Here += 2
                            End If
                        End If

                    Case Else
                        If IsFullWidth(ch) Then
                            ch = MakeHalfWidth(ch)
                            GoTo FullWidthRepeat2
                        End If
                End Select
            End If

            ' ####################################################
            ' //  Produce a value for the literal.
            ' ####################################################

            Dim IntegralValue As UInt64
            Dim FloatingValue As Double
            Dim DecimalValue As Decimal
            Dim Overflows As Boolean = False

            If literalKind = NumericLiteralKind.Integral Then
                If IntegerLiteralStart = IntegerLiteralEnd Then
                    Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_Syntax)
                Else
                    IntegralValue = 0

                    If Base = LiteralBase.Decimal Then
                        ' Init For loop
                        For LiteralCharacter As Integer = IntegerLiteralStart To IntegerLiteralEnd - 1
                            Dim LiteralCharacterValue As Char = Peek(LiteralCharacter)
                            If LiteralCharacterValue = "_"c Then
                                Continue For
                            End If
                            Dim NextCharacterValue As UInteger = IntegralLiteralCharacterValue(LiteralCharacterValue)

                            If IntegralValue < 1844674407370955161UL OrElse
                              (IntegralValue = 1844674407370955161UL AndAlso NextCharacterValue <= 5UI) Then

                                IntegralValue = (IntegralValue * 10UL) + NextCharacterValue
                            Else
                                Overflows = True
                                Exit For
                            End If
                        Next

                        If TypeCharacter <> TypeCharacter.ULongLiteral AndAlso IntegralValue > Long.MaxValue Then
                            Overflows = True
                        End If
                    Else
                        Dim Shift As Integer = If(Base = LiteralBase.Hexadecimal, 4, If(Base = LiteralBase.Octal, 3, 1))
                        Dim OverflowMask As UInt64 = If(Base = LiteralBase.Hexadecimal, &HF000000000000000UL, If(Base = LiteralBase.Octal, &HE000000000000000UL, &H8000000000000000UL))

                        ' Init For loop
                        For LiteralCharacter As Integer = IntegerLiteralStart To IntegerLiteralEnd - 1
                            Dim LiteralCharacterValue As Char = Peek(LiteralCharacter)
                            If LiteralCharacterValue = "_"c Then
                                Continue For
                            End If

                            If (IntegralValue And OverflowMask) <> 0 Then
                                Overflows = True
                            End If

                            IntegralValue = (IntegralValue << Shift) + IntegralLiteralCharacterValue(LiteralCharacterValue)
                        Next
                    End If

                    If TypeCharacter = TypeCharacter.None Then
                        ' nothing to do
                    ElseIf TypeCharacter = TypeCharacter.Integer OrElse TypeCharacter = TypeCharacter.IntegerLiteral Then
                        If (Base = LiteralBase.Decimal AndAlso IntegralValue > &H7FFFFFFF) OrElse
                            IntegralValue > &HFFFFFFFFUI Then

                            Overflows = True
                        End If

                    ElseIf TypeCharacter = TypeCharacter.UIntegerLiteral Then
                        If IntegralValue > &HFFFFFFFFUI Then
                            Overflows = True
                        End If

                    ElseIf TypeCharacter = TypeCharacter.ShortLiteral Then
                        If (Base = LiteralBase.Decimal AndAlso IntegralValue > &H7FFF) OrElse
                            IntegralValue > &HFFFF Then

                            Overflows = True
                        End If

                    ElseIf TypeCharacter = TypeCharacter.UShortLiteral Then
                        If IntegralValue > &HFFFF Then
                            Overflows = True
                        End If

                    Else
                        Debug.Assert(TypeCharacter = TypeCharacter.Long OrElse
                                 TypeCharacter = TypeCharacter.LongLiteral OrElse
                                 TypeCharacter = TypeCharacter.ULongLiteral,
                        "Integral literal value computation is lost.")
                    End If
                End If

            Else
                ' // Copy the text of the literal to deal with fullwidth
                Dim scratch = GetScratch()
                For i = 0 To literalWithoutTypeChar - 1
                    Dim curCh = Peek(i)
                    If curCh <> "_"c Then
                        scratch.Append(If(IsFullWidth(curCh), MakeHalfWidth(curCh), curCh))
                    End If
                Next
                Dim LiteralSpelling = GetScratchTextInterned(scratch)

                If literalKind = NumericLiteralKind.Decimal Then
                    ' Attempt to convert to Decimal.
                    Overflows = Not GetDecimalValue(LiteralSpelling, DecimalValue)
                Else
                    If TypeCharacter = TypeCharacter.Single OrElse TypeCharacter = TypeCharacter.SingleLiteral Then
                        ' // Attempt to convert to single
                        Dim SingleValue As Single
                        If Not RealParser.TryParseFloat(LiteralSpelling, SingleValue) Then
                            Overflows = True
                        Else
                            FloatingValue = SingleValue
                        End If
                    Else
                        ' // Attempt to convert to double.
                        If Not RealParser.TryParseDouble(LiteralSpelling, FloatingValue) Then
                            Overflows = True
                        End If
                    End If
                End If
            End If

            Dim result As SyntaxToken
            Select Case literalKind
                Case NumericLiteralKind.Integral
                    result = MakeIntegerLiteralToken(precedingTrivia, Base, TypeCharacter, If(Overflows Or UnderscoreInWrongPlace, 0UL, IntegralValue), Here)
                Case NumericLiteralKind.Float
                    result = MakeFloatingLiteralToken(precedingTrivia, TypeCharacter, If(Overflows Or UnderscoreInWrongPlace, 0.0F, FloatingValue), Here)
                Case NumericLiteralKind.Decimal
                    result = MakeDecimalLiteralToken(precedingTrivia, TypeCharacter, If(Overflows Or UnderscoreInWrongPlace, 0D, DecimalValue), Here)
                Case Else
                    Throw ExceptionUtilities.UnexpectedValue(literalKind)
            End Select

            If Overflows Then
                result = DirectCast(result.AddError(ErrorFactory.ErrorInfo(ERRID.ERR_Overflow)), SyntaxToken)
            End If

            If UnderscoreInWrongPlace Then
                result = DirectCast(result.AddError(ErrorFactory.ErrorInfo(ERRID.ERR_Syntax)), SyntaxToken)
            ElseIf LeadingUnderscoreUsed Then
                result = CheckFeatureAvailability(result, Feature.DigitSeparators, Me.Options)
            ElseIf UnderscoreUsed Then
                result = CheckFeatureAvailability(result, Feature.DigitSeparators, Me.Options)
            End If

            If Base = LiteralBase.Binary Then
                result = CheckFeatureAvailability(result, Feature.BinaryLiterals, Me.Options)
            End If

            Return result
        End Function

        Private Shared Function GetDecimalValue(text As String, <Out()> ByRef value As Decimal) As Boolean

            ' Use Decimal.TryParse to parse value. Note: the behavior of
            ' Decimal.TryParse differs from Dev11 in the following cases:
            '
            ' 1. [-]0eNd where N > 0
            '     The native compiler ignores sign and scale and treats such cases
            '     as 0e0d. Decimal.TryParse fails so these cases are compile errors.
            '     [Bug #568475]
            ' 2. Decimals with significant digits below 1e-49
            '     The native compiler considers digits below 1e-49 when rounding.
            '     Decimal.TryParse ignores digits below 1e-49 when rounding. This
            '     difference is perhaps the most significant since existing code will
            '     continue to compile but constant values may be rounded differently.
            '     [Bug #568494]

            Return Decimal.TryParse(text, NumberStyles.AllowDecimalPoint Or NumberStyles.AllowExponent, CultureInfo.InvariantCulture, value)
        End Function

        Private Function ScanIntLiteral(
               ByRef ReturnValue As Integer,
               ByRef Here As Integer
           ) As Boolean
            Debug.Assert(Here >= 0)
            Dim ch As Char = Nothing
            If Not TryGet(ch, Here) OrElse Not IsDecimalDigit(ch) Then Return False

            Dim IntegralValue As Integer = IntegralLiteralCharacterValue(ch)
            Here += 1

            While TryGet(ch, Here) AndAlso IsDecimalDigit(ch)

                Dim nextDigit = IntegralLiteralCharacterValue(ch)
                If IntegralValue < 214748364 OrElse
                    (IntegralValue = 214748364 AndAlso nextDigit < 8) Then

                    IntegralValue = IntegralValue * 10 + nextDigit
                    Here += 1
                Else
                    Return False
                End If
            End While

            ReturnValue = IntegralValue
            Return True
        End Function

        Private Function ScanDateLiteral(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Debug.Assert(CanGet)
            Debug.Assert(IsHash(Peek()))
            Dim ch AS Char = Nothing
            Dim Here As Integer = 1 'skip #
            Dim FirstValue As Integer
            Dim TimePart As (HaveTime As Boolean, HaveMinute As Boolean, HaveSeconds As Boolean, HaveAM As Boolean, HavePM As Boolean, HH AS Int32, MM AS Int32, SS As Int32) = (False, False, False, False, False, 0, 0, 0)
            Dim datePart As (HaveDate As Boolean, HaveYear As Boolean, IsYearFirst As Boolean, Is2Digit As Boolean, YYYY As Int32, MM As Int32, DD AS Int32)

            Dim DateIsInvalid As Boolean = False
            Dim badDate As SyntaxToken = Nothing


            ' // Unfortunately, we can't fall back on OLE Automation's date parsing because
            ' // they don't have the same range as the URT's DateTime class

            ' // First, eat any whitespace
            Here = GetWhitespaceLength(Here)

            Dim FirstValueStart As Integer = Here

            ' // The first thing has to be an integer, although it's not clear what it is yet
            If Not ScanIntLiteral(FirstValue, Here) Then
                Return Nothing

            End If

            ' // If we see a /, then it's a date
            If TryGet(ch, Here) AndAlso IsDateSeparatorCharacter(ch) Then
                If Not TryScan_DatePart(precedingTrivia, ch, here, firstValue, FirstValueStart, datePart, badDate) Then Return badDate
            End If

            If Not TryScan_TimePart(precedingTrivia, ch, Here, FirstValue, TimePart, datePart, badDate) Then REturn badDate

            If Not TryGet(ch, Here) OrElse Not IsHash(ch) Then
                IsABadDate(precedingTrivia, here, badDate) : Return badDate
            End If

            Here += 1
            Validate_DatePart(DateIsInvalid, datePart)

            Validate_TimePart(TimePart, DateIsInvalid)

            ' // Ok, we've got a valid value. Now make into an i8.

            If Not DateIsInvalid Then
                Dim DateTimeValue As New DateTime(datePart.YYYY, datePart.MM, datePart.DD, TimePart.HH, TimePart.MM, TimePart.SS)
                Dim result = MakeDateLiteralToken(precedingTrivia, DateTimeValue, Here)

                If datePart.IsYearFirst Then
                    result = CheckFeatureAvailability(result, Feature.YearFirstDateLiterals, Options)
                End If

                Return result
            Else
                Return MakeBadToken(precedingTrivia, Here, ERRID.ERR_InvalidDate)
            End If
        End Function


        Private Function TryScan_DatePart(
                                           precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                           ch As Char,
                                     ByRef here As Int32,
                                           firstValue As Int32,
                                           FirstValueStart As int32,
                                     ByRef datePart As (HaveDate As Boolean, HaveYear As Boolean, IsYearFirst As Boolean, Is2Digit As Boolean, YYYY As Int32, MM As Int32, DD AS Int32),
                                     ByRef badDate As SyntaxToken
                                         ) As Boolean
            Dim FirstDateSeparator As Char = ch
            ' // We've got a date
            datePart.HaveDate = True
            Here += 1

            ' Is the first value a year?
            ' It is a year if it consists of exactly 4 digits.
            ' Condition below uses 5 because we already skipped the separator.
            If Here - FirstValueStart = 5 Then
                datePart.HaveYear = True
                datePart.IsYearFirst = True
                datePart.YYYY = FirstValue

                ' // We have to have a month value
                If Not ScanIntLiteral(datePart.MM, Here) Then Return IsABadDate(precedingTrivia, here, BadDate)

                ' Do we have a day value?
                If TryGet(ch, Here) AndAlso IsDateSeparatorCharacter(ch) Then
                    ' // Check to see they used a consistent separator

                    If ch <> FirstDateSeparator Then Return IsABadDate(precedingTrivia, here, badDate)

                    ' // Yes.
                    Here += 1

                    If Not ScanIntLiteral(datePart.DD, Here) Then Return IsABadDate(precedingTrivia, here, badDate)
                End If
            Else
                ' First value is month
                datePart.MM = FirstValue

                ' // We have to have a day value

                If Not ScanIntLiteral(datePart.DD, Here) Then Return IsABadDate(precedingTrivia, here, badDate)

                ' // Do we have a year value?

                If TryGet(ch, Here) AndAlso IsDateSeparatorCharacter(ch) Then
                    ' // Check to see they used a consistent separator

                    If ch <> FirstDateSeparator Then Return IsABadDate(precedingTrivia, here, badDate)

                    ' // Yes.
                    datePart.HaveYear = True
                    Here += 1

                    Dim YearStart As Integer = Here

                    If Not ScanIntLiteral(datePart.YYYY, Here) Then Return IsABadDate(precedingTrivia, here, badDate)

                    If (Here - YearStart) = 2 Then
                       datePart.Is2Digit = True
                    End If
                End If
            End If

            Here = GetWhitespaceLength(Here)
            return True
        End Function

        Private Shared ReadOnly DaysToMonth365 As Int32() = {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365}
        Private Shared ReadOnly DaysToMonth366 As Int32() = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}


        Private Sub Validate_DatePart _ 
            (
        ByRef DateIsInvalid As Boolean,
        ByRef datePart As (HaveDate As Boolean, HaveYear As Boolean, IsYearFirst As Boolean, Is2Digit As Boolean, YYYY As Integer, MM As Integer, DD As Integer)
            )
          Dim DaysToMonth As Integer() = Nothing
          With datePart 
            ' // OK, now we've got all the values, let's see if we've got a valid date
            If .HaveDate Then
                If Not .MM.IsBetween(1, 12) Then DateIsInvalid = True
                ' // We'll check Days in a moment...

                If Not .HaveYear Then
                    DateIsInvalid = True
                    .YYYY = 1
                End If

                ' // Check if not a leap year

                If Not ((.YYYY Mod 4 = 0) AndAlso (Not (.YYYY Mod 100 = 0) OrElse (.YYYY Mod 400 = 0))) Then
                    DaysToMonth = DaysToMonth365
                Else
                    DaysToMonth = DaysToMonth366
                End If

                If Not DateIsInvalid AndAlso Not .DD.IsBetween(1, DaysToMonth(.MM) - DaysToMonth(.MM - 1)) Then DateIsInvalid = True
                If .Is2Digit Then DateIsInvalid = True
                If Not .YYYY.IsBetween(1, 9999) Then DateIsInvalid = True
            Else
                .MM = 1
                .DD = 1
                .YYYY = 1
                DaysToMonth = DaysToMonth365
            End If
          End With
        End Sub

        Private Function TryScan_TimePart(
                                       precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                 ByRef ch As Char,
                                 ByRef Here As Integer,
                                       FirstValue As Integer,
                                 ByRef TimePart As (HaveTime As Boolean, HaveMinute As Boolean, HaveSeconds As Boolean, HaveAM As Boolean, HavePM As Boolean, HH As Integer, MM As Integer, SS As Integer),
                                 ByRef datePart As (HaveDate As Boolean, HaveYear As Boolean, IsYearFirst As Boolean, Is2Digit As Boolean, YYYY As Integer, MM As Integer, DD As Integer),
                                 ByRef badDate As SyntaxToken
                                     ) As Boolean
            ' // If we haven't seen a date, assume it's a time value

            If Not datePart.HaveDate Then
                TimePart.HaveTime = True
                TimePart.HH = FirstValue
            Else
                ' // We did see a date. See if we see a time value...

                If ScanIntLiteral(TimePart.HH, Here) Then
                    ' // Yup.
                    TimePart.HaveTime = True
                End If
            End If

            If TimePart.HaveTime Then
                ' // Do we see a :?

                If TryGet(ch, Here) AndAlso IsColon(ch) Then
                    Here += 1

                    ' // Now let's get the minute value

                    If Not ScanIntLiteral(TimePart.MM, Here) Then Return IsABadDate(precedingTrivia, here, badDate)

                    TimePart.HaveMinute = True

                    ' // Do we have a second value?

                    If TryGet(ch, Here) AndAlso IsColon(ch) Then
                        ' // Yes.
                        TimePart.HaveSeconds = True
                        Here += 1

                        If Not ScanIntLiteral(TimePart.SS, Here) Then Return IsABadDate(precedingTrivia, here, badDate)
                    End If
                End If

                Here = GetWhitespaceLength(Here)

                ' // Check AM/PM

                If TryGet(ch, Here) Then
                    If ch.IsEither("A"c, FULLWIDTH_LATIN_CAPITAL_LETTER_A,
                                   "a"c, FULLWIDTH_LATIN_SMALL_LETTER_A) Then

                        TimePart.HaveAM = True
                        Here += 1

                    ElseIf ch.IsEither("P"c, FULLWIDTH_LATIN_CAPITAL_LETTER_P,
                                       "p"c, FULLWIDTH_LATIN_SMALL_LETTER_P) Then

                        TimePart.HavePM = True
                        Here += 1

                    End If

                    If TryGet(ch, Here) AndAlso (TimePart.HaveAM OrElse TimePart.HavePM) Then
                        If ch.IsEither("M"c, FULLWIDTH_LATIN_CAPITAL_LETTER_M,
                                       "m"c, FULLWIDTH_LATIN_SMALL_LETTER_M) Then

                            Here = GetWhitespaceLength(Here + 1)

                        Else
                            Return IsABadDate(precedingTrivia, here, badDate)
                        End If
                    End If
                End If

                ' // If there's no minute/second value and no AM/PM, it's invalid

                If Not TimePart.HaveMinute AndAlso Not TimePart.HaveAM AndAlso Not TimePart.HavePM Then Return IsABadDate(precedingTrivia, here, BadDate)
            End If
            return true
        End Function

        Private Shared Sub Validate_TimePart _ 
            (
        ByRef TimePart As (HaveTime As Boolean, HaveMinute As Boolean, HaveSeconds As Boolean, HaveAM As Boolean, HavePM As Boolean, HH As Integer, MM As Integer, SS As Integer),
        ByRef DateIsInvalid As Boolean
            )

          With TimePart
            If .HaveTime Then
                If .HaveAM OrElse .HavePM Then
                    ' // 12-hour value
                    If Not  .HH.IsBetween(1,12) Then DateIsInvalid = True
                    If .HaveAM Then
                        .HH = .HH Mod 12
                    ElseIf .HavePM Then
                        .HH = .HH + 12
                        If .HH = 24 Then .HH = 12
                    End If

                Else
                    If Not .HH.IsBetween(0, 23) Then DateIsInvalid = True
                End If

                If .HaveMinute Then
                    If Not .MM.IsBetween(0, 59) Then DateIsInvalid = True
                Else
                    .MM = 0
                End If

                If .HaveSeconds Then
                    If Not .SS.IsBetween(0, 59) Then DateIsInvalid = True
                Else
                    TimePart.SS = 0
                End If
            Else
                .HH = 0
                .MM = 0
                .SS = 0
            End If
          End With
        End Sub


        Private Function IsABadDate(
                                         precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                   byref here As Int32,
                                   byref badDate As SyntaxToken
                                       ) As Boolean
            ' // If we can find a closing #, then assume it's a malformed date,
            ' // otherwise, it's not a date
            Dim ch As Char = Nothing
            While TryGet(ch, Here)
                If IsHash(ch) OrElse IsNewLine(ch) Then
                    Exit While
                End If
                Here += 1
            End While
            If Not TryGet(ch, Here) OrElse IsNewLine(ch) Then
                ' // No closing #
                badDate = Nothing
            Else
                Debug.Assert(IsHash(ch))
                Here += 1  ' consume trailing #
                badDate = MakeBadToken(precedingTrivia, Here, ERRID.ERR_InvalidDate)
            End If
            Return False
        End Function

        Private Function ScanStringLiteral(precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)) As SyntaxToken
            Debug.Assert(CanGet)
            Debug.Assert(IsDoubleQuote(Peek))

            Dim length As Integer = 1
            Dim ch As Char
            Dim followingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)

            ' // Check for a Char literal, which can be of the form:
            ' // """"c or "<anycharacter-except-">"c

            If CanGet(3) AndAlso IsDoubleQuote(Peek(2)) Then
                If IsDoubleQuote(Peek(1)) Then
                    If IsDoubleQuote(Peek(3)) AndAlso
                       CanGet(4) AndAlso
                       IsLetterC(Peek(4)) Then

                        ' // Double-quote Char literal: """"c
                        Return MakeCharacterLiteralToken(precedingTrivia, """"c, 5)
                    End If

                ElseIf IsLetterC(Peek(3)) Then
                    ' // Char literal.  "x"c
                    Return MakeCharacterLiteralToken(precedingTrivia, Peek(1), 4)
                End If
            End If

            If TryGet(ch, 2) AndAlso IsDoubleQuote(Peek(1)) AndAlso IsLetterC(ch) Then
                ' // Error. ""c is not a legal char constant
                Return MakeBadToken(precedingTrivia, 3, ERRID.ERR_IllegalCharConstant)
            End If

            Dim haveNewLine As Boolean = False

            Dim scratch = GetScratch()
            While TryGet(ch, length)
                If IsDoubleQuote(ch) Then
                    If TryGet(ch, length + 1) Then
                        If IsDoubleQuote(ch) Then
                            ' // An escaped double quote
                            scratch.Append(""""c)
                            length += 2
                            Continue While
                        Else
                            ' // The end of the char literal.
                            If IsLetterC(ch) Then
                                ' // Error. "aad"c is not a legal char constant

                                ' // +2 to include both " and c in the token span
                                scratch.Clear()
                                Return MakeBadToken(precedingTrivia, length + 2, ERRID.ERR_IllegalCharConstant)
                            End If
                        End If
                    End If

                    ' the double quote was a valid string terminator.
                    length += 1
                    Dim spelling = GetTextNotInterned(length)
                    followingTrivia = ScanSingleLineTrivia()

                    ' NATURAL TEXT, NO INTERNING
                    Dim result As SyntaxToken = SyntaxFactory.StringLiteralToken(spelling, GetScratchText(scratch), precedingTrivia.Node, followingTrivia.Node)

                    If haveNewLine Then
                        result = CheckFeatureAvailability(result, Feature.MultilineStringLiterals, Options)
                    End If

                    Return result

                ElseIf IsNewLine(ch) Then
                    If _isScanningDirective Then
                        Exit While
                    End If

                    haveNewLine = True
                End If

                scratch.Append(ch)
                length += 1
            End While

            ' CC has trouble to prove this after the loop
            Debug.Assert(CanGet(length - 1))

            '// The literal does not have an explicit termination.
            ' DIFFERENT: here in IDE we used to report string token marked as unterminated

            Dim sp = GetTextNotInterned(length)
            followingTrivia = ScanSingleLineTrivia()
            Dim strTk = SyntaxFactory.StringLiteralToken(sp, GetScratchText(scratch), precedingTrivia.Node, followingTrivia.Node)
            Dim StrTkErr = strTk.SetDiagnostics({ErrorFactory.ErrorInfo(ERRID.ERR_UnterminatedStringLiteral)})

            Debug.Assert(StrTkErr IsNot Nothing)
            Return DirectCast(StrTkErr, SyntaxToken)
        End Function

        Friend Shared Function TryIdentifierAsContextualKeyword(id As IdentifierTokenSyntax, ByRef k As SyntaxKind) As Boolean
            Debug.Assert(id IsNot Nothing)

            If id.PossibleKeywordKind <> SyntaxKind.IdentifierToken Then
                k = id.PossibleKeywordKind
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Try to convert an Identifier to a Keyword.  Called by the parser when it wants to force
        ''' an identifier to be a keyword.
        ''' </summary>
        Friend Function TryIdentifierAsContextualKeyword(id As IdentifierTokenSyntax, ByRef k As KeywordSyntax) As Boolean
            Debug.Assert(id IsNot Nothing)

            Dim kind As SyntaxKind = SyntaxKind.IdentifierToken
            If TryIdentifierAsContextualKeyword(id, kind) Then
                k = MakeKeyword(id)
                Return True
            End If

            Return False
        End Function

        Friend Function TryTokenAsContextualKeyword(t As SyntaxToken, ByRef k As KeywordSyntax) As Boolean
            If t Is Nothing Then
                Return False
            End If

            If t.Kind = SyntaxKind.IdentifierToken Then
                Return TryIdentifierAsContextualKeyword(DirectCast(t, IdentifierTokenSyntax), k)
            End If

            Return False
        End Function

        Friend Shared Function TryTokenAsKeyword(t As SyntaxToken, ByRef kind As SyntaxKind) As Boolean

            If t Is Nothing Then
                Return False
            End If

            If t.IsKeyword Then
                kind = t.Kind
                Return True
            End If

            If t.Kind = SyntaxKind.IdentifierToken Then
                Return TryIdentifierAsContextualKeyword(DirectCast(t, IdentifierTokenSyntax), kind)
            End If

            Return False
        End Function

        Friend Shared Function IsContextualKeyword(t As SyntaxToken, ParamArray kinds As SyntaxKind()) As Boolean
            Dim kind As SyntaxKind = Nothing
            Return If( TryTokenAsKeyword(t, kind), Array.IndexOf(kinds, kind) >= 0, False)
        End Function

        Private Function IsIdentifierStartCharacter(c As Char) As Boolean
            Return (_isScanningForExpressionCompiler AndAlso c = "$"c) OrElse SyntaxFacts.IsIdentifierStartCharacter(c)
        End Function

    End Class
End Namespace

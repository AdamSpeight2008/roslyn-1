' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------

Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.Text
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Parser

        Friend Function ParseParameterList() As ParameterListSyntax
            If CurrentToken.Kind = SyntaxKind.OpenParenToken Then

                Dim openParen As PunctuationSyntax = Nothing
                Dim closeParen As PunctuationSyntax = Nothing
                Dim parameters = ParseParameters(openParen, closeParen)

                Return SyntaxFactory.ParameterList(openParen, parameters, closeParen)
            Else
                Return Nothing
            End If
        End Function

        ' /*********************************************************************
        ' *
        ' * Function:
        ' *     Parser::ParseParameters
        ' *
        ' * Purpose:
        ' *     Parses a parenthesized parameter list of non-optional followed by
        ' *     optional parameters (if any).
        ' *
        ' **********************************************************************/
        '
        ' File:Parser.cpp
        ' Lines: 9598 - 9598
        ' ParameterList* .Parser::ParseParameters( [ _Inout_ bool& ErrorInConstruct ] [ _Out_ Token*& openParen ] [ _Out_ Token*& closeParen ] )
        '
        Private Function ParseParameters(
                                    ByRef openParen As PunctuationSyntax,
                                    ByRef closeParen As PunctuationSyntax
                                        ) As CoreInternalSyntax.SeparatedSyntaxList(Of ParameterSyntax)
            Debug.Assert(CurrentToken.Kind = SyntaxKind.OpenParenToken, CalledOnWrongToken(NameOf(ParseParameters)))
            TryGetTokenAndEatNewLine(SyntaxKind.OpenParenToken, openParen)

            Dim parameters = _pool.AllocateSeparated(Of ParameterSyntax)()

            If CurrentToken.Kind <> SyntaxKind.CloseParenToken Then

                ' Loop through the list of parameters.

                Do
                    Dim attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax) = Nothing

                    If Me.CurrentToken.Kind = Global.Microsoft.CodeAnalysis.VisualBasic.SyntaxKind.LessThanToken Then
                        attributes = ParseAttributeLists(False)
                    End If

                    Dim paramSpecifiers As ParameterSpecifiers = 0
                    Dim modifiers = ParseParameterSpecifiers(paramSpecifiers)
                    Dim param = ParseParameter(attributes, modifiers)

                    ' TODO - Bug 889301 - Dev10 does a resync here when there is an error.  That prevents ERRID_InvalidParameterSyntax below from
                    ' being reported. For now keep backwards compatibility.
                    If param.ContainsDiagnostics Then
                        param = param.AddTrailingSyntax(ResyncAt({SyntaxKind.CommaToken, SyntaxKind.CloseParenToken}))
                    End If

                    Dim comma As PunctuationSyntax = Nothing
                    If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then

                        If CurrentToken.Kind <> SyntaxKind.CloseParenToken AndAlso Not MustEndStatement(CurrentToken) Then

                            ' Check the ')' on the next line
                            If IsContinuableEOL() Then
                                If PeekToken(1).Kind = SyntaxKind.CloseParenToken Then
                                    parameters.Add(param)
                                    Exit Do
                                End If
                            End If

                            param = param.AddTrailingSyntax(ResyncAt({SyntaxKind.CommaToken, SyntaxKind.CloseParenToken}), ERRID.ERR_InvalidParameterSyntax)

                            If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then
                                parameters.Add(param)
                                Exit Do
                            End If

                        Else
                            parameters.Add(param)
                            Exit Do

                        End If
                    End If

                    parameters.Add(param)
                    parameters.AddSeparator(comma)
                Loop

            End If

            ' Current token is left at either tkRParen, EOS

            TryEatNewLineAndGetToken(SyntaxKind.CloseParenToken, closeParen, createIfMissing:=True)

            Dim result = parameters.ToList()

            _pool.Free(parameters)

            Return result

        End Function

        ' /*********************************************************************
        ' *
        ' * Function:
        ' *     Parser::ParseParameterSpecifiers
        ' *
        ' * Purpose:
        ' *
        ' **********************************************************************/
        '
        ' File:Parser.cpp
        ' Lines: 9748 - 9748
        ' ParameterSpecifierList* .Parser::ParseParameterSpecifiers( [ _Inout_ bool& ErrorInConstruct ] )
        '
        Private Function ParseParameterSpecifiers(
                                             ByRef specifiers As ParameterSpecifiers
                                                 ) As CoreInternalSyntax.SyntaxList(Of KeywordSyntax)
            Dim keywords = Me._pool.Allocate(Of KeywordSyntax)()

            specifiers = 0

            'TODO - Move these checks to Binder_Utils.DecodeParameterModifiers  

            Do
                Dim specifier As ParameterSpecifiers
                Dim keyword As KeywordSyntax

                Select Case (CurrentToken.Kind)

                    Case SyntaxKind.ByValKeyword
                        keyword = DirectCast(CurrentToken, KeywordSyntax)
                        If (specifiers And ParameterSpecifiers.ByRef) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_MultipleParameterSpecifiers)
                        End If
                        specifier = ParameterSpecifiers.ByVal

                    Case SyntaxKind.ByRefKeyword
                        keyword = DirectCast(CurrentToken, KeywordSyntax)
                        If (specifiers And ParameterSpecifiers.ByVal) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_MultipleParameterSpecifiers)

                        ElseIf (specifiers And ParameterSpecifiers.ParamArray) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_ParamArrayMustBeByVal)
                        End If
                        specifier = ParameterSpecifiers.ByRef

                    Case SyntaxKind.OptionalKeyword
                        keyword = DirectCast(CurrentToken, KeywordSyntax)
                        If (specifiers And ParameterSpecifiers.ParamArray) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_MultipleOptionalParameterSpecifiers)
                        End If
                        specifier = ParameterSpecifiers.Optional

                    Case SyntaxKind.ParamArrayKeyword
                        keyword = DirectCast(CurrentToken, KeywordSyntax)
                        If (specifiers And ParameterSpecifiers.Optional) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_MultipleOptionalParameterSpecifiers)
                        ElseIf (specifiers And ParameterSpecifiers.ByRef) <> 0 Then
                            keyword = ReportSyntaxError(keyword, ERRID.ERR_ParamArrayMustBeByVal)
                        End If
                        specifier = ParameterSpecifiers.ParamArray

                    Case Else
                        Dim result = keywords.ToList
                        Me._pool.Free(keywords)

                        Return result
                End Select

                If (specifiers And specifier) <> 0 Then
                    keyword = ReportSyntaxError(keyword, ERRID.ERR_DuplicateParameterSpecifier)
                Else
                    specifiers = specifiers Or specifier
                End If

                keywords.Add(keyword)

                GetNextToken()
            Loop
        End Function

        ''' <summary>
        '''     Parameter -> Attributes? ParameterModifiers* ParameterIdentifier ("as" TypeName)? ("=" ConstantExpression)?
        ''' </summary>
        ''' <param name="attributes"></param>
        ''' <param name="modifiers"></param>
        ''' <returns></returns>
        ''' <remarks>>This replaces both ParseParameter and ParseOptionalParameter in Dev10</remarks>
        Private Function ParseParameter(
                                         attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax),
                                         modifiers As CoreInternalSyntax.SyntaxList(Of KeywordSyntax)
                                       ) As ParameterSyntax

            Dim paramName = ParseModifiedIdentifier(False, False)

            If paramName.ContainsDiagnostics Then

                ' If we see As before a comma or RParen, then assume that
                ' we are still on the same parameter. Otherwise, don't resync
                ' and allow the caller to decide how to recover.

                If PeekAheadFor(SyntaxKind.AsKeyword, SyntaxKind.CommaToken, SyntaxKind.CloseParenToken) = SyntaxKind.AsKeyword Then
                    paramName = ResyncAt(paramName, SyntaxKind.AsKeyword)
                End If
            End If

            Dim optionalAsClause As SimpleAsClauseSyntax = Nothing
            Dim asKeyword As KeywordSyntax = Nothing

            If TryGetToken(SyntaxKind.AsKeyword, asKeyword) Then
                Dim typeName = ParseGeneralType()

                optionalAsClause = SyntaxFactory.SimpleAsClause(asKeyword, Nothing, typeName)

                If optionalAsClause.ContainsDiagnostics Then
                    optionalAsClause = ResyncAt(optionalAsClause, SyntaxKind.EqualsToken, SyntaxKind.CommaToken, SyntaxKind.CloseParenToken)
                End If

            End If
            Dim initializer As EqualsValueSyntax = Nothing

            If TryParseEqualsExpression(initializer) Then

                '' TODO - Move these errors (ERRID.ERR_DefaultValueForNonOptionalParamout, ERRID.ERR_ObsoleteOptionalWithoutValue) of the parser. 
                '' These are semantic errors. The grammar allows the syntax. 
                'If Not (modifiers.Any AndAlso modifiers.Any(SyntaxKind.OptionalKeyword)) Then
                '    initializer = SyntaxFactory.EqualsValue(initializer.EqualsToken.AddError(ERRID.ERR_DefaultValueForNonOptionalParam), initializer.Value)
                'ElseIf modifiers.Any AndAlso modifiers.Any(SyntaxKind.OptionalKeyword) Then
                '    initializer = SyntaxFactory.EqualsValue(
                'ReportSyntaxError(InternalSyntaxFactory.MissingPunctuation(SyntaxKind.EqualsToken), ERRID.ERR_ObsoleteOptionalWithoutValue),
                'initializer.Value)
                'End If


                If initializer.Value IsNot Nothing Then

                    If initializer.Value.ContainsDiagnostics Then
                        initializer = SyntaxFactory.EqualsValue(initializer.EqualsToken,
                                                                ResyncAt(initializer.Value, SyntaxKind.CommaToken, SyntaxKind.CloseParenToken))
                    End If
                End If
            Else

            End If
            Return SyntaxFactory.Parameter(attributes, modifiers, paramName, optionalAsClause, initializer)
        End Function


    End Class

End Namespace

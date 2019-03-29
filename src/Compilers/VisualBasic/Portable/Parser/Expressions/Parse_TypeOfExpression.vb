' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'
'============ Methods for parsing portions of executable statements ==
'
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.VisualBasic.LanguageFeatures.CheckFeatureAvailability

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Friend Partial Class Parser

        ''' <summary>
        ''' Parse TypeOf ... Is ... or TypeOf ... IsNot ...
        ''' TypeOfExpression -> "TypeOf" Expression "Is|IsNot" LineTerminator? TypeName
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ParseTypeOf() As TypeOfBaseExpressionSyntax
            Debug.Assert(CurrentToken.Kind = SyntaxKind.TypeOfKeyword, "must be at TypeOf.")
            Dim [typeOf] As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)

            ' Consume 'TypeOf'.
            GetNextToken()

            Dim exp As ExpressionSyntax = ParseExpressionCore(OperatorPrecedence.PrecedenceRelational) 'Dev10 uses ParseVariable
            If exp.ContainsDiagnostics Then exp = ResyncAt(exp, SyntaxKind.IsKeyword, SyntaxKind.IsNotKeyword)

            Dim operatorToken As KeywordSyntax = Nothing

            Dim current As SyntaxToken = CurrentToken
            Select Case Current.Kind
                   Case SyntaxKind.IsKeyword
                        operatorToken = DirectCast(current, KeywordSyntax)
                        GetNextToken()
                        TryEatNewLine(ScannerState.VB)

                   Case SyntaxKind.IsNotKeyword
                        operatorToken = DirectCast(current, KeywordSyntax).CheckFeatureAvailability(Feature.TypeOfIsNot, Options)
                        GetNextToken()
                        TryEatNewLine(ScannerState.VB)

                   Case Else
                        operatorToken = DirectCast(HandleUnexpectedToken(SyntaxKind.IsKeyword), KeywordSyntax)
                        operatortoken = ReportSyntaxError(operatorToken, ERRID.ERR_MissingIsInTypeOf)
            End Select

            current = CurrentToken
            If current.Kind = SyntaxKind.OpenBraceToken Then Return Parse_RestAs_TypeOfMany([typeOf], exp, operatorToken, current)
            Return Parse_RestAs_TypeOfSingle([typeOf], exp, operatorToken)
        End Function

        Private Function Parse_RestAs_TypeOfSingle _
            (
                [typeOf]        As KeywordSyntax,
                exp             As ExpressionSyntax,
                operatorToken   As KeywordSyntax
            ) As TypeOfBaseExpressionSyntax

            Dim typeName = ParseGeneralType()
            Dim kind     = If(operatorToken.Kind = SyntaxKind.IsNotKeyword,
                              SyntaxKind.TypeOfIsNotExpression, SyntaxKind.TypeOfIsExpression)
            Return SyntaxFactory.TypeOfExpression(kind, [typeOf], exp, operatorToken, typeName)
        End Function

        Private Function Parse_RestAs_TypeOfMany _ 
            (
                [typeOf]        As KeywordSyntax,
                exp             As ExpressionSyntax,
                operatorToken   As KeywordSyntax,
          ByRef current         As SyntaxToken
            ) As TypeOfBaseExpressionSyntax

            Dim openingBrace As PunctuationSyntax = Nothing
            Dim closingBrace As PunctuationSyntax = Nothing
            Dim types = ParseTypeList(openingBrace, closingBrace)
            current = current.CheckFeatureAvailability(Feature.TypeOfMany, Options)
            Dim mkind As SyntaxKind = If(operatorToken.Kind = SyntaxKind.IsNotKeyword,
                                         SyntaxKind.TypeOfManyIsNotExpression, SyntaxKind.TypeOfManyIsExpression)
            Return SyntaxFactory.TypeOfManyExpression(mkind, [typeOf], exp, operatorToken, openingBrace, types, closingBrace)
        End Function

        Private Function ParseTypeList _ 
            (
        ByRef openingBrace  As PunctuationSyntax,
        ByRef closingBrace  As PunctuationSyntax
            ) As CoreInternalSyntax.SeparatedSyntaxList(Of TypeSyntax)

          Debug.Assert(CurrentToken.Kind = SyntaxKind.OpenBraceToken, $"{NameOf(ParseTypeList)} list parsing confused.")

          TryGetTokenAndEatNewLine(SyntaxKind.OpenBraceToken, openingBrace, True)

          Dim typelist = _pool.AllocateSeparated(Of TypeSyntax)()

          If CurrentToken.Kind <> SyntaxKind.CloseBraceToken Then
             ' Loop through the list of parameters.
             Do
               Dim typeName = ParseGeneralType()
               If typeName.ContainsDiagnostics Then typeName = typeName.AddTrailingSyntax(ResyncAt({SyntaxKind.CommaToken, SyntaxKind.CloseBraceToken}))

               Dim comma As PunctuationSyntax = Nothing
               If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then
                  If CurrentToken.Kind <> SyntaxKind.CloseBraceToken AndAlso Not MustEndStatement(CurrentToken) Then
                     ' Check the '}' on the next line
                     If IsContinuableEOL() AndAlso PeekToken(1).Kind = SyntaxKind.CloseBraceToken Then typelist.Add(typeName) : Exit Do
 
                     typeName = typeName.AddTrailingSyntax(ResyncAt({SyntaxKind.CommaToken, SyntaxKind.CloseBraceToken}), ERRID.ERR_InvalidTypeSyntax)

                     If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then typelist.Add(typeName) : Exit Do

                  Else
                     typelist.Add(typeName)
                     Exit Do
                  End If
               End If

               typelist.Add(typeName).AddSeparator(comma)
             Loop
          End If

          ' Current token is left at either tkRParen, EOS

          Dim ok = TryEatNewLineAndGetToken(SyntaxKind.CloseBraceToken, closingBrace, createIfMissing:=True)
          Return typelist.ToListAndFree(_pool)
        End Function

    End Class

End Namespace

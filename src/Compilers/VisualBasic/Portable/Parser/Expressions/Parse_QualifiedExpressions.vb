' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'
'============ Methods for parsing portions of executable statements ==
'
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Features.CheckFeatureAvailability

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

  Friend Partial Class Parser

    ' /*********************************************************************
    ' *
    ' * Function:
    ' *     Parser::ParseQualifiedExpr
    ' *
    ' * Purpose:
    ' *     Parses a dot or bang reference, starting at the dot or bang.
    ' *
    ' **********************************************************************/
    ' [in] token starting term
    ' [in] stuff before "." or "!"
    '
    ' File: Parser.cpp
    ' Lines: 16211 - 16211
    ' Expression* .Parser::ParseQualifiedExpr( [ _In_ Token* Start ] [ _In_opt_ ParseTree::Expression* Term ] [ _Inout_ bool& ErrorInConstruct ] )
    Private Function ParseQualifiedExpr _ 
            (
              Term As ExpressionSyntax
            ) As ExpressionSyntax

     Debug.Assert(CurrentToken.Kind.IsEither(SyntaxKind.DotToken, SyntaxKind.ExclamationToken),
                  "Must be on either a '.' or '!' when entering parseQualifiedExpr()")

     Dim DotOrBangToken As PunctuationSyntax = DirectCast(CurrentToken, PunctuationSyntax)

     Dim prevPrevToken = PrevToken
     GetNextToken()

     Dim output As ExpressionSyntax = Nothing

     If TryParseFlagEnumExpr_Or_QualifiedExpr(Term, DotOrBangToken, output) Then Return output

     If (CurrentToken.IsEndOfLine() AndAlso Not CurrentToken.IsEndOfParse()) Then

        Debug.Assert(CurrentToken.Kind = SyntaxKind.StatementTerminatorToken AndAlso
                   PrevToken.Kind = SyntaxKind.DotToken,
                   "We shouldn't get here without .<eol> tokens")

        '/* We know we are sitting on an EOL preceded by a tkDot.  What we need to catch is the
        '   case where a tkDot following an EOL isn't preceded by a valid token.  Bug Dev10_429652  For example:
        '   with i <eol>
        '     .  <-- this is bad.  This . follows an EOL and isn't preceded by a tkID.  Can't have it dangling like this
        '     field = 42
        '*/
        If (prevPrevToken Is Nothing OrElse
            prevPrevToken.Kind = SyntaxKind.StatementTerminatorToken) Then
           ' if ( CurrentToken->m_Prev->m_Prev == NULL || // make sure we can look back far enough.  We know we can look back once, but twice we need to test
           '     CurrentToken->m_Prev->m_Prev->m_TokenType == tkEOL ) // Make sure there is something besides air before the '.' DEV10_486908

           Dim missingIdent = ReportSyntaxError(InternalSyntaxFactory.MissingIdentifier, ERRID.ERR_ExpectedIdentifier)

           ' We are sitting on the tkEOL so let's just return this and keep parsing.  No ReSync() needed here, in other words.
           Return SyntaxFactory.SimpleMemberAccessExpression(Term, DotOrBangToken, SyntaxFactory.IdentifierName(missingIdent))

        ElseIf Not NextLineStartsWithStatementTerminator() Then
           '//ILC: undone
           '//       Right now we don't continue after a "." when the following tokens indicate XML member access
           '//       We should probably enable this.
           TryEatNewLineIfNotFollowedBy(SyntaxKind.DotToken)
        End If
     End If

     ' Decide whether we're parsing:
     ' 1. Element axis i.e. ".<ident>"
     ' 2. Attribute axis i.e. ".@ident" or ".@<ident>
     ' 3. Descendant axis i.e. "...<ident>"
     ' 4. Regular CLR member axis i.e. ".ident"
     Select Case (CurrentToken.Kind)
            Case SyntaxKind.AtToken
                 Return ParseAs_XmlAttributeAccessExpression(term, DotOrBangToken)

            Case SyntaxKind.LessThanToken
                 Return ParseAs_XmlElementAccessExpression(term, DotOrBangToken)

            Case SyntaxKind.DotToken
                 If PeekToken(1).Kind = SyntaxKind.DotToken Then
                    ' Consume the 2nd and 3rd dots and remember that this is descendant axis
                    Dim secondDotToken = DirectCast(CurrentToken, PunctuationSyntax)
                    GetNextToken()
                    Dim thirdDotToken As PunctuationSyntax = Nothing
                    TryGetToken(SyntaxKind.DotToken, thirdDotToken)
                    ' Parse the Xml element name
                    TryEatNewLineIfFollowedBy(SyntaxKind.LessThanToken)
                    Dim name As XmlBracketedNameSyntax
                    If CurrentToken.Kind = SyntaxKind.LessThanToken Then
                       name = ParseBracketedXmlQualifiedName()
                    Else
                       name = ReportExpectedXmlBracketedName()
                    End If
                    Return SyntaxFactory.XmlMemberAccessExpression(SyntaxKind.XmlDescendantAccessExpression, Term, DotOrBangToken, secondDotToken, thirdDotToken, name)
                 End If

            Case Else
                 Dim name = ParseSimpleNameExpressionAllowingKeywordAndTypeArguments()
                 Return SyntaxFactory.SimpleMemberAccessExpression(Term, DotOrBangToken, name)
     End Select
     
     'This is reachable with the following invalid syntax.
     '    p.  
     '      x
     ' 
     ' or 
     '   p..y
     Dim result As ExpressionSyntax
     If CurrentToken.Kind = SyntaxKind.AtToken Then
        Dim missingName = DirectCast(InternalSyntaxFactory.MissingToken(SyntaxKind.XmlNameToken), XmlNameTokenSyntax)
        result = SyntaxFactory.XmlMemberAccessExpression(SyntaxKind.XmlAttributeAccessExpression,
                                                   Term,
                                                   DotOrBangToken, Nothing, Nothing,
                                                   ReportSyntaxError(InternalSyntaxFactory.XmlName(Nothing, missingName), ERRID.ERR_ExpectedXmlName))
     Else
         result = SyntaxFactory.SimpleMemberAccessExpression(Term, DotOrBangToken, ReportSyntaxError(InternalSyntaxFactory.IdentifierName(InternalSyntaxFactory.MissingIdentifier), ERRID.ERR_ExpectedIdentifier))
     End If

     Return result
    End Function

    Private Function ParseAs_XmlElementAccessExpression(Term As ExpressionSyntax, DotOrBangToken As PunctuationSyntax) As XmlMemberAccessExpressionSyntax
      ' Remember that this is element axis
      ' Parse the Xml element name
      Dim name = ParseBracketedXmlQualifiedName()
      Return SyntaxFactory.XmlMemberAccessExpression(SyntaxKind.XmlElementAccessExpression, Term, DotOrBangToken, Nothing, Nothing, name)
    End Function


    Private Function ParseAs_XmlAttributeAccessExpression(Term As ExpressionSyntax, DotOrBangToken As PunctuationSyntax) As XmlMemberAccessExpressionSyntax
      Dim atToken = DirectCast(CurrentToken, PunctuationSyntax)
      Dim name As XmlNodeSyntax

      ' Do not accept space or anything else between @ and name or <
      If atToken.HasTrailingTrivia Then
         GetNextToken(ScannerState.VB)
         atToken = ReportSyntaxError(atToken, ERRID.ERR_ExpectedXmlName)
         atToken = atToken.AddTrailingSyntax(ResyncAt())
         name = SyntaxFactory.XmlName(Nothing, DirectCast(InternalSyntaxFactory.MissingToken(SyntaxKind.XmlNameToken), XmlNameTokenSyntax))

      Else
         ' Parse the Xml attribute name (allow name with and without angle brackets)
         If PeekNextToken(ScannerState.VB).Kind = SyntaxKind.LessThanToken Then
            ' Consume the @ and remember that this is an element axis
            GetNextToken(ScannerState.Element)
            name = ParseBracketedXmlQualifiedName()
         Else
            ' Consume the @ and remember that this is attribute axis
            GetNextToken(ScannerState.VB)
            name = ParseXmlQualifiedNameVB()

            If name.HasLeadingTrivia Then
               atToken = ReportSyntaxError(atToken, ERRID.ERR_ExpectedXmlName)
               atToken.AddTrailingSyntax(name)
               atToken.AddTrailingSyntax(ResyncAt())
               name = SyntaxFactory.XmlName(Nothing, DirectCast(InternalSyntaxFactory.MissingToken(SyntaxKind.XmlNameToken), XmlNameTokenSyntax))
            End If
         End If
      End If
      Return SyntaxFactory.XmlMemberAccessExpression(SyntaxKind.XmlAttributeAccessExpression, Term, DotOrBangToken, atToken, Nothing, name)
     End Function

  End Class

End Namespace

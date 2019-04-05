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

        Private Function ParseFlagsEnumExpr _
            (
              term  As ExpressionSyntax,
              op    As FlagsEnumOperatorSyntax
            ) As ExpressionSyntax

          Dim prevPrevToken = PrevToken
          GetNextToken()
          Select Case CurrentToken.Kind
                 Case SyntaxKind.IdentifierToken : Return ParseFlagsEnumExpr_WithIdentifier(term, op)
                 Case SyntaxKind.OpenParenToken  : Return ParseFlagsEnumExpr_WithParenthesisedExpression(term, op)
          End Select
          Dim expr = ParseExpression()
          Return SyntaxFactory.FlagsEnumOperationExpression(term, op, expr)
        End Function

        Private Function ParseFlagsEnumExpr_WithIdentifier _
            (
              term  As ExpressionSyntax,
              op    As FlagsEnumOperatorSyntax
            ) As ExpressioNSyntax
          ' Term FlagsEnumOper Identifier
          Dim Name = ParseIdentifierNameAllowingKeyword(True)
          If Name IsNot Nothing Then Return SyntaxFactory.FlagsEnumOperationExpression(term, op, Name)
          Return SyntaxFactory.FlagsEnumOperationExpression(term, op, Name.AddError(ERRID.ERR_ExpectedIdentifier))
        End Function

        Private Function ParseFlagsEnumExpr_WithParenthesisedExpression _
            (
              term  As ExpressionSyntax,
              op    As FlagsEnumOperatorSyntax
            ) As ExpressionSyntax
          ' Term FlagsEnumOper ( ParenthesizedExpression | TupleLiteral )
          Dim pexpr As ExpressionSyntax = ParseParenthesizedExpressionOrTupleLiteral()
          If pexpr Is Nothing Then Return SyntaxFactory.FlagsEnumOperationExpression(term, op, pexpr.AddError(ERRID.ERR_ExpectedExpression))
          ' ( ParenthesizedExpression | TupleLiteral )
          If pexpr.Kind <> SyntaxKind.ParenthesizedExpression Then pexpr = AddError(pexpr, ERRID.ERR_ExpectedExpression)
          ' ParenthesisedExpresssion
          Return SyntaxFactory.FlagsEnumOperationExpression(term, op, pexpr)
        End function

        Private Function TryParseFlagEnumExpr_Or_QualifiedExpr _
            (
              term      As ExpressionSyntax,
              op        As PunctuationSyntax,
  <Out> ByRef output    As ExpressionSyntax
            ) As Boolean
          output = Nothing
          If op.Kind <> SyntaxKind.ExclamationToken Then Return False
          If PeekToken(0).Kind = SyntaxKind.IdentifierToken Then
             Dim name = ParseIdentifierNameAllowingKeyword()
             output = SyntaxFactory.DictionaryAccessExpression(term, op, name)
          Else
             Dim fop = SyntaxFactory.FlagsEnumIsSetToken(op.Text, op.GetLeadingTrivia, op.GetTrailingTrivia)
             Dim expr = ParseExpression()
             output = SyntaxFactory.FlagsEnumOperationExpression(term, fop, expr)
          End If
          Return output IsNot Nothing
        End Function



    End Class

End Namespace

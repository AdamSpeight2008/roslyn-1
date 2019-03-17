' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Parser
'-----------------------------------------------------------------------------
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

  Partial Friend Class Parser

    Private Function ParseCaseStatement() As CaseStatementSyntax
      Debug.Assert(CurrentToken.Kind = SyntaxKind.CaseKeyword, NameOf(ParseCaseStatement) + " called on wrong token.")

      Dim caseKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)
      GetNextToken()

      Dim caseClauses = _pool.AllocateSeparated(Of CaseClauseSyntax)()
      Dim elseKeyword As KeywordSyntax = Nothing

      If CurrentToken.Kind = SyntaxKind.ElseKeyword Then
          elseKeyword = DirectCast(CurrentToken, KeywordSyntax)
          GetNextToken() '// get off ELSE

          Dim caseClause = SyntaxFactory.ElseCaseClause(elseKeyword)
          caseClauses.Add(caseClause)
      Else

          ParseCaseClauseList(caseClauses)
      End If

      Dim separatedCaseClauses = caseClauses.ToList()
      _pool.Free(caseClauses)

      Dim statement As CaseStatementSyntax

      If elseKeyword Is Nothing Then
          statement = SyntaxFactory.CaseStatement(caseKeyword, separatedCaseClauses)
      Else
          statement = SyntaxFactory.CaseElseStatement(caseKeyword, separatedCaseClauses)
      End If

      'If CurrentToken.Kind = SyntaxKind.WhenKeyword Then

      'End If
      Return statement
    End Function

    Private Sub ParseCaseClauseList(caseClauses As SeparatedSyntaxListBuilder(Of InternalSyntax.CaseClauseSyntax))
      Do
      Loop Until TryParse_CaseClause(caseClauses)
    End Sub

    Private Function TryParse_CaseClause(caseClauses As  SeparatedSyntaxListBuilder(Of CaseClauseSyntax)) As Boolean
        Dim StartCase  As SyntaxKind = CurrentToken.Kind ' dev10_500588 Snap the start of the expression token AFTER we've moved off the EOL (if one is present)
        Dim caseClause As CaseClauseSyntax = nothing

        If StartCase = SyntaxKind.IsKeyword OrElse SyntaxFacts.IsRelationalOperator(StartCase) Then
           caseClause = ParseCaseClause_Relational()
        else
           caseClause = ParsePossibleRangeCaseClause()
        End If
        caseClauses.Add(caseClause)
      REturn TryParse_Comma(caseClauses)
    End Function

    Private Function TryParse_Comma(of NodeType As GreenNode)(commaSeparated As SeparatedSyntaxListBuilder(Of NodeType)) As Boolean
      Dim comma As PunctuationSyntax = Nothing
      If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then Return False
      commaSeparated.AddSeparator(comma)
      Return True
    End Function

    Private Function ParsePossibleRangeCaseClause() As CaseClauseSyntax
      Dim caseClause As CaseClauseSyntax
      Dim value As ExpressionSyntax = ParseExpressionCore()
      If value.ContainsDiagnostics Then value = ResyncAt(value, SyntaxKind.ToKeyword)
      Dim toKeyword As KeywordSyntax = Nothing
      If TryGetToken(SyntaxKind.ToKeyword, toKeyword) Then
         Dim upperBound As ExpressionSyntax = ParseExpressionCore()
         If upperBound.ContainsDiagnostics Then upperBound = ResyncAt(upperBound)
         caseClause = SyntaxFactory.RangeCaseClause(value, toKeyword, upperBound)
      Else
         caseClause = SyntaxFactory.SimpleCaseClause(value)
      End If
      Return caseClause
    End Function

    Private Function ParseCaseClause_Relational() As CaseClauseSyntax
      Dim caseClause As CaseClauseSyntax
      ' dev10_526560 Allow implicit newline after IS
      Dim optionalIsKeyword As KeywordSyntax = Nothing
      TryGetTokenAndEatNewLine(SyntaxKind.IsKeyword, optionalIsKeyword)

      If SyntaxFacts.IsRelationalOperator(CurrentToken.Kind) Then

        Dim relationalOperator = DirectCast(CurrentToken, PunctuationSyntax)
        GetNextToken() ' get off relational operator
        TryEatNewLine() ' dev10_503248

        Dim CaseExpr As ExpressionSyntax = ParseExpressionCore()
        If CaseExpr.ContainsDiagnostics Then CaseExpr = ResyncAt(CaseExpr)
        caseClause = SyntaxFactory.RelationalCaseClause(RelationalOperatorKindToCaseKind(relationalOperator.Kind), optionalIsKeyword, relationalOperator, CaseExpr)

      Else
        ' Since we saw IS, create a relational case.
        ' This helps intellisense do a drop down of
        ' the operators that can follow "Is".
        Dim relationalOperator = ReportSyntaxError(InternalSyntaxFactory.MissingPunctuation(SyntaxKind.EqualsToken), ERRID.ERR_ExpectedRelational)
        
        caseClause = ResyncAt(InternalSyntaxFactory.RelationalCaseClause(SyntaxKind.CaseEqualsClause, optionalIsKeyword, relationalOperator, InternalSyntaxFactory.MissingExpression))
      End If

      Return caseClause
    End Function

    Private Shared Function RelationalOperatorKindToCaseKind(kind As SyntaxKind) As SyntaxKind
      Select Case kind
             Case SyntaxKind.LessThanToken              : Return SyntaxKind.CaseLessThanClause
             Case SyntaxKind.LessThanEqualsToken        : Return SyntaxKind.CaseLessThanOrEqualClause
             Case SyntaxKind.EqualsToken                : Return SyntaxKind.CaseEqualsClause
             Case SyntaxKind.LessThanGreaterThanToken   : Return SyntaxKind.CaseNotEqualsClause
             Case SyntaxKind.GreaterThanToken           : Return SyntaxKind.CaseGreaterThanClause
             Case SyntaxKind.GreaterThanEqualsToken     : Return SyntaxKind.CaseGreaterThanOrEqualClause
      End Select
      Debug.Assert(False, "Wrong relational operator kind")
      Return Nothing
    End Function

    Private Function ParseSelectStatement() As SelectStatementSyntax
       Debug.Assert(CurrentToken.Kind = SyntaxKind.SelectKeyword, "ParseSelectStatement called on wrong token.")
      Dim selectKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)
      GetNextToken() ' get off SELECT
      ' Allow the expected CASE token to be present or not.
      Dim optionalCaseKeyword As KeywordSyntax = Nothing
      TryGetToken(SyntaxKind.CaseKeyword, optionalCaseKeyword)
      Dim value As ExpressionSyntax = ParseExpressionCore()
      If value.ContainsDiagnostics Then            value = ResyncAt(value)

      Dim statement = SyntaxFactory.SelectStatement(selectKeyword, optionalCaseKeyword, value)
      Return statement
    End Function

  End Class

End Namespace

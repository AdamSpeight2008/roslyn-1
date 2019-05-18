' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Parser
'-----------------------------------------------------------------------------
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Parser

        Private Function CalledWrongToken(tokenName As String) As String
            tokenName = If(tokenName, String.Empty)
            Return $"{TokenName} called on wrong token"
        End Function

        Private Function ParseUsingStatement() As UsingStatementSyntax
            Debug.Assert(CurrentToken.Kind = SyntaxKind.UsingKeyword, CalledWrongToken(NameOf(ParseUsingStatement)))

            Dim usingKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)
            GetNextToken()

            Dim optionalExpression As ExpressionSyntax = Nothing
            Dim variables As CodeAnalysis.Syntax.InternalSyntax.SeparatedSyntaxList(Of VariableDeclaratorSyntax) = Nothing
            Dim optionalWithKeyword As KeywordSyntax = Nothing

            If CurrentToken.Kind = SyntaxKind.WithKeyword Then
                optionalWithKeyword = CheckFeatureAvailability(Feature.UsingWithStatement, DirectCast(CurrentToken, KeywordSyntax))
                GetNextToken() ' Get off the With
            End If

            Dim nextToken As SyntaxToken = PeekToken(1)

            ' change from Dev10: allowing as new with multiple variable names, e.g. "Using a, b As New C1()"
            If nextToken.Kind = SyntaxKind.AsKeyword OrElse
               nextToken.Kind = SyntaxKind.EqualsToken OrElse
               nextToken.Kind = SyntaxKind.CommaToken OrElse
               nextToken.Kind = SyntaxKind.QuestionToken Then

                variables = ParseVariableDeclaration(allowAsNewWith:=True)
            Else
                optionalExpression = ParseExpressionCore()
            End If

            If optionalWithKeyword IsNot Nothing AndAlso optionalExpression Is Nothing AndAlso variables.Count <> 1 Then
               optionalWithKeyword = ReportSyntaxError(optionalWithKeyword, ERRID.ERR_InvalidUsingWithStatement)
            End If

            'TODO - not resyncing here may cause errors to differ from Dev10.

            'No need to resync on error.  This will be handled by GetStatementTerminator

            Dim statement = SyntaxFactory.UsingStatement(usingKeyword, optionalWithKeyword, optionalExpression, variables)

            Return statement
        End Function

    End Class

End Namespace

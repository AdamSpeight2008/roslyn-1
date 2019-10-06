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

        Private Function CalledOnWrongToken(name As String) As String
            name = If(name, String.Empty)
            Return name & " called on the wrong token."
        End Function

        Private Function IsStartOfPossibleGeneric() As Boolean
            Return CurrentToken.Kind = SyntaxKind.OpenParenToken
        End Function

        Private Function TryParseSimpleAsClause(
                                           ByRef asClause As SimpleAsClauseSyntax,
                                        Optional checkForTypeAttributes As Boolean = False
                                               ) As Boolean
            ' eg As type
            asClause = Nothing
            If CurrentToken.Kind = SyntaxKind.AsKeyword Then
                Dim asKeyword = DirectCast(CurrentToken, KeywordSyntax)
                GetNextToken()
                Dim typeAttributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax) = Nothing
                If checkForTypeAttributes Then

                    If CurrentToken.Kind = SyntaxKind.LessThanToken Then
                        typeAttributes = ParseAttributeLists(False)
                    End If
                End If

                Dim type = ParseGeneralType()

                If type.ContainsDiagnostics Then
                    type = ResyncAt(type)
                End If

                asClause = SyntaxFactory.SimpleAsClause(asKeyword, typeAttributes, type)
            End If
            Return asClause IsNot Nothing
        End Function

        Private Function TryParseEqualsExpression(
                                             ByRef equalsValue As EqualsValueSyntax
                                                 ) As Boolean
            Dim optionalEquals As PunctuationSyntax = Nothing
            Dim expr As ExpressionSyntax = Nothing

            If TryGetTokenAndEatNewLine(SyntaxKind.EqualsToken, optionalEquals) Then

                expr = ParseExpressionCore()

                If expr.ContainsDiagnostics Then
                    ' Resync at EOS so we don't get any more errors.
                    expr = ResyncAt(expr)
                End If

                equalsValue = SyntaxFactory.EqualsValue(optionalEquals, expr)
            Else
                equalsValue = Nothing
            End If
            Return equalsValue IsNot Nothing
        End Function

    End Class

End Namespace

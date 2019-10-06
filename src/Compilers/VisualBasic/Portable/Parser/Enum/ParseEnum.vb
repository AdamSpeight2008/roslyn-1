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

        ' /*********************************************************************
        ' *
        ' * Function:
        ' *     Parser::ParseEnumStatement
        ' *
        ' * Purpose:
        ' *     Parses: Enum <ident>
        ' *
        ' **********************************************************************/
        ' File:Parser.cpp
        ' Lines: 4352 - 4352
        ' EnumTypeStatement* .Parser::ParseEnumStatement( [ ParseTree::AttributeSpecifierList* Attributes ] [ ParseTree::SpecifierList* Specifiers ] [ Token* Start ] [ _Inout_ bool& ErrorInConstruct ] )
        Private Function ParseEnumStatement(
                                    Optional attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax) = Nothing,
                                    Optional modifiers As CoreInternalSyntax.SyntaxList(Of KeywordSyntax) = Nothing
                                           ) As EnumStatementSyntax

            Debug.Assert(CurrentToken.Kind = SyntaxKind.EnumKeyword, CalledOnWrongToken(NameOf(ParseEnumStatement)))

            Dim enumKeyword = DirectCast(CurrentToken, KeywordSyntax)
            Dim optionalUnderlyingType As AsClauseSyntax = Nothing

            GetNextToken() ' Get Off ENUM

            Dim identifier = ParseIdentifier()

            If IsStartOfPossibleGeneric() Then
                ' Enums cannot be generic
                Dim genericParameters = ReportSyntaxError(ParseGenericParameters, ERRID.ERR_GenericParamsOnInvalidMember)
                identifier = identifier.AddTrailingSyntax(genericParameters)
            End If

            If identifier.ContainsDiagnostics Then
                identifier = identifier.AddTrailingSyntax(ResyncAt({SyntaxKind.AsKeyword}))
            End If

            TryParseSimpleAsClause(optionalUnderlyingType)

            Return SyntaxFactory.EnumStatement(attributes, modifiers, enumKeyword, identifier, optionalUnderlyingType)
        End Function

        ' /*********************************************************************
        ' *
        ' * Function:
        ' *     Parser::ParseEnumMember
        ' *
        ' * Purpose:
        ' *     Parses an enum member definition.
        ' *
        ' *     Does NOT advance to next line so caller can recover from errors.
        ' *
        ' **********************************************************************/
        '
        ' File:Parser.cpp
        ' Lines: 4438 - 4438
        ' Statement* .Parser::ParseEnumMember( [ _Inout_ bool& ErrorInConstruct ] )
        Private Function ParseEnumMemberOrLabel(attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax)) As StatementSyntax

            If Not attributes.Any() AndAlso ShouldParseAsLabel() Then
                Return ParseLabel()
            End If

            ' The current token should be an Identifier
            ' The Dev10 code used to look ahead to see if the statement was a declaration to exit out of en enum declaration.
            ' The new parser calls ParseEnumMember from ParseDeclaration so this look ahead is not necessary.  The enum block
            ' parsing will terminate when the bad statement is added to the enum block context.

            ' Check to see if this construct is a valid module-level declaration.
            ' If it is, end the current enum context and reparse the statement.
            ' (This case is important for automatic end insertion.)

            Dim ident As IdentifierTokenSyntax = ParseIdentifier()

            If ident.ContainsDiagnostics Then
                ident = ident.AddTrailingSyntax(ResyncAt({SyntaxKind.EqualsToken}))
            End If

            ' See if there is an expression

            Dim initializer As EqualsValueSyntax = Nothing

            TryParseEqualsExpression(initializer)

            Dim statement As EnumMemberDeclarationSyntax = SyntaxFactory.EnumMemberDeclaration(attributes, ident, initializer)

            Return statement

        End Function


        Private Function TryParseEqualsExpression(ByRef equalsExpression As EqualsValueSyntax) As Boolean
            Dim eqaulsSign As PunctuationSyntax = Nothing
            If TryGetTokenAndEatNewLine(SyntaxKind.EqualsToken, eqaulsSign) Then

                Dim expr = ParseExpressionCore()

                If expr.ContainsDiagnostics Then
                    ' Resync at EOS so we don't get any more errors.
                    expr = ResyncAt(expr)
                End If

                equalsExpression = SyntaxFactory.EqualsValue(eqaulsSign, expr)
            Else
                equalsExpression = Nothing
            End If

            Return equalsExpression IsNot Nothing
        End Function

#Region "Helpers"

        Private Function CalledOnWrongToken(name As String) As String
            name = If(name, String.Empty)
            Return name & " called on the wrong token."
        End Function

        Private Function IsStartOfPossibleGeneric() As Boolean
            Return CurrentToken.Kind = SyntaxKind.OpenParenToken
        End Function

        Private Function TryParseSimpleAsClause(ByRef asClause As AsClauseSyntax) As Boolean
            ' eg As type
            asClause = Nothing
            If CurrentToken.Kind = SyntaxKind.AsKeyword Then
                Dim asKeyword = DirectCast(CurrentToken, KeywordSyntax)

                GetNextToken() ' get off AS

                Dim typeName = ParseTypeName()

                If typeName.ContainsDiagnostics Then
                    typeName = typeName.AddTrailingSyntax(ResyncAt())
                End If

                asClause = SyntaxFactory.SimpleAsClause(asKeyword, Nothing, typeName)
            End If
            Return asClause IsNot Nothing
        End Function

#End Region

    End Class

End Namespace

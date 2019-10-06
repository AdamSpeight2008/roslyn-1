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
            Dim optionalUnderlyingType As SimpleAsClauseSyntax = Nothing

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


    End Class

End Namespace

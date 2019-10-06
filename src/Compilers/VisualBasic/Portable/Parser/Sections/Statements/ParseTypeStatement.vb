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
        ' *     Parser::ParseTypeStatement
        ' *
        ' * Purpose:
        ' *
        ' **********************************************************************/
        '
        ' [in] specifiers on decl
        ' [in] token starting Enum statement
        ' File:Parser.cpp
        ' Lines: 4563 - 4563
        ' TypeStatement* .Parser::ParseTypeStatement( [ ParseTree::AttributeSpecifierList* Attributes ] [ ParseTree::SpecifierList* Specifiers ] [ Token* Start ] [ _Inout_ bool& ErrorInConstruct ] )
        Private Function ParseTypeStatement(
                                    Optional attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax) = Nothing,
                                    Optional modifiers As CoreInternalSyntax.SyntaxList(Of KeywordSyntax) = Nothing
                                           ) As TypeStatementSyntax

            Dim kind As SyntaxKind
            Dim optionalTypeParameters As TypeParameterListSyntax = Nothing

            Dim typeKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)
            GetNextToken()

            Select Case (typeKeyword.Kind)

                Case SyntaxKind.ModuleKeyword
                    kind = SyntaxKind.ModuleStatement

                Case SyntaxKind.ClassKeyword
                    kind = SyntaxKind.ClassStatement

                Case SyntaxKind.StructureKeyword
                    kind = SyntaxKind.StructureStatement

                Case SyntaxKind.InterfaceKeyword
                    kind = SyntaxKind.InterfaceStatement

                Case Else
                    Throw ExceptionUtilities.UnexpectedValue(typeKeyword.Kind)
            End Select

            Dim ident As IdentifierTokenSyntax = ParseIdentifier()

            If ident.ContainsDiagnostics Then
                ident = ident.AddTrailingSyntax(ResyncAt({SyntaxKind.OfKeyword, SyntaxKind.OpenParenToken}))
            End If

            If IsStartOfPossibleGeneric() Then
                ' Modules cannot be generic
                '
                If kind = SyntaxKind.ModuleStatement Then
                    ident = ident.AddTrailingSyntax(ReportGenericParamsDisallowedError(ERRID.ERR_ModulesCannotBeGeneric))
                Else
                    optionalTypeParameters = ParseGenericParameters()
                End If
            End If

            Dim statement As TypeStatementSyntax = InternalSyntaxFactory.TypeStatement(kind, attributes, modifiers, typeKeyword, ident, optionalTypeParameters)

            If (kind = SyntaxKind.ModuleStatement OrElse kind = SyntaxKind.InterfaceStatement) AndAlso statement.Modifiers.Any(SyntaxKind.PartialKeyword) Then
                statement = CheckFeatureAvailability(If(kind = SyntaxKind.ModuleStatement, Feature.PartialModules, Feature.PartialInterfaces), statement)
            End If

            Return statement
        End Function


    End Class

End Namespace

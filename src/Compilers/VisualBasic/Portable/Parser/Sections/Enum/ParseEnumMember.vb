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
        ' *     Parser::ParseEnumMember
        ' *
        ' * Purpose:
        ' *     Parses an enum member definition.
        ' *
        ' *     Does NOT advance to next line so caller can recover from errors.
        ' *
        ' **********************************************************************/

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
            If TryParseEqualsExpression(initializer) Then
            Else
            End If

            Dim statement As EnumMemberDeclarationSyntax = SyntaxFactory.EnumMemberDeclaration(attributes, ident, initializer)

            Return statement

        End Function

    End Class

End Namespace

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
        ' *     Parser::ParseSubDeclaration
        ' *
        ' * Purpose:
        ' *
        ' **********************************************************************/
        '
        ' [in] specifiers on definition
        ' [in] token starting definition
        '
        ' File: Parser.cpp
        ' Lines: 8358 - 8358
        ' MethodDeclarationStatement* .Parser::ParseSubDeclaration( [ ParseTree::AttributeSpecifierList* Attributes ] [ ParseTree::SpecifierList* Specifiers ] [ _In_ Token* Start ] [ bool IsDelegate ] [ _Inout_ bool& ErrorInConstruct ] )
        '
        Private Function ParseSubStatement(
                                            attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax),
                                            modifiers As CoreInternalSyntax.SyntaxList(Of KeywordSyntax)
                                          ) As MethodBaseSyntax

            Dim subKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)

            Debug.Assert(subKeyword.Kind = SyntaxKind.SubKeyword, CalledOnWrongToken(NameOf(ParseSubStatement)))

            GetNextToken()

            Dim save_isInMethodDeclarationHeader As Boolean = _isInMethodDeclarationHeader
            _isInMethodDeclarationHeader = True

            Dim save_isInAsyncMethodDeclarationHeader As Boolean = _isInAsyncMethodDeclarationHeader
            Dim save_isInIteratorMethodDeclarationHeader As Boolean = _isInIteratorMethodDeclarationHeader

            _isInAsyncMethodDeclarationHeader = modifiers.Any(SyntaxKind.AsyncKeyword)
            _isInIteratorMethodDeclarationHeader = modifiers.Any(SyntaxKind.IteratorKeyword)

            Dim newKeyword As KeywordSyntax = Nothing
            Dim name As IdentifierTokenSyntax = Nothing
            Dim genericParams As TypeParameterListSyntax = Nothing
            Dim parameters As ParameterListSyntax = Nothing
            Dim handlesClause As HandlesClauseSyntax = Nothing
            Dim implementsClause As ImplementsClauseSyntax = Nothing

            ' Dev10_504604 we are parsing a method declaration and will need to let the scanner know that we
            ' are so the scanner can correctly identify attributes vs. xml while scanning the declaration.

            'davidsch - It is not longer necessary to force the scanner state here.  The scanner will only scan xml when the parser explicitly tells it to scan xml.

            ' Nodekind.NewKeyword is allowed as a Sub name but no other keywords.
            If CurrentToken.Kind = SyntaxKind.NewKeyword Then
                newKeyword = DirectCast(CurrentToken, KeywordSyntax)
                GetNextToken()
            End If

            ParseSubOrDelegateStatement(If(newKeyword Is Nothing, SyntaxKind.SubStatement, SyntaxKind.SubNewStatement), name, genericParams, parameters, handlesClause, implementsClause)

            ' We should be at the end of the statement.
            _isInMethodDeclarationHeader = save_isInMethodDeclarationHeader
            _isInAsyncMethodDeclarationHeader = save_isInAsyncMethodDeclarationHeader
            _isInIteratorMethodDeclarationHeader = save_isInIteratorMethodDeclarationHeader

            'Create the Sub declaration
            If newKeyword Is Nothing Then
                Return SyntaxFactory.SubStatement(attributes, modifiers, subKeyword, name, genericParams, parameters, Nothing, handlesClause, implementsClause)
            Else
                If handlesClause IsNot Nothing Then
                    newKeyword = newKeyword.AddError(ERRID.ERR_NewCannotHandleEvents) ' error should be on "New"
                End If

                If implementsClause IsNot Nothing Then
                    newKeyword = newKeyword.AddError(ERRID.ERR_ImplementsOnNew) ' error should be on "New"
                End If

                If genericParams IsNot Nothing Then
                    newKeyword = newKeyword.AddTrailingSyntax(genericParams)
                End If

                Dim ctorDecl = SyntaxFactory.SubNewStatement(attributes, modifiers, subKeyword, newKeyword, parameters)

                ' do not forget unexpected handles and implements even if unexpected
                ctorDecl = ctorDecl.AddTrailingSyntax(handlesClause)
                ctorDecl = ctorDecl.AddTrailingSyntax(implementsClause)

                Return ctorDecl
            End If

        End Function


        Private Sub ParseSubOrDelegateStatement(
                                                 kind As SyntaxKind,
                                           ByRef ident As IdentifierTokenSyntax,
                                           ByRef optionalGenericParams As TypeParameterListSyntax,
                                           ByRef optionalParameters As ParameterListSyntax,
                                           ByRef handlesClause As HandlesClauseSyntax,
                                           ByRef implementsClause As ImplementsClauseSyntax
                                               )

            Debug.Assert(kind = SyntaxKind.SubStatement OrElse
                         kind = SyntaxKind.SubNewStatement OrElse
                         kind = SyntaxKind.DelegateSubStatement, CalledOnWrongToken(NameOf(ParseSubOrDelegateStatement)))

            'The current token is on the Sub or Delegate's name

            ' Parse the name only for Delegates and Subs.  Constructors have already grabbed the New keyword.
            If kind <> SyntaxKind.SubNewStatement Then
                ident = ParseIdentifier()

                If ident.ContainsDiagnostics Then
                    ident = ident.AddTrailingSyntax(ResyncAt({SyntaxKind.OpenParenToken, SyntaxKind.OfKeyword}))
                End If
            End If

            ' Dev10_504604 we are parsing a method declaration and will need to let the scanner know that we
            ' are so the scanner can correctly identify attributes vs. xml while scanning the declaration.

            If BeginsGeneric() Then
                If kind = SyntaxKind.SubNewStatement Then

                    ' We want to do this error checking here during parsing and not in
                    ' declared (which would have been more ideal) because for the invalid
                    ' case, when this error occurs, we don't want any parse errors for
                    ' parameters to show up.

                    ' We want other errors such as those on regular parameters reported too,
                    ' so don't mark ErrorInConstruct, but instead use a temp.
                    '
                    optionalGenericParams = ReportGenericParamsDisallowedError(ERRID.ERR_GenericParamsOnInvalidMember)
                Else
                    optionalGenericParams = ParseGenericParameters()
                End If
            End If

            optionalParameters = ParseParameterList()

            ' See if we have the HANDLES or the IMPLEMENTS clause on this procedure.

            If CurrentToken.Kind = SyntaxKind.HandlesKeyword Then
                handlesClause = ParseHandlesList()

                If kind = SyntaxKind.DelegateSubStatement Then
                    ' davidsch - This error was reported in Declared in Dev10
                    handlesClause = ReportSyntaxError(handlesClause, ERRID.ERR_DelegateCantHandleEvents)
                End If
            ElseIf CurrentToken.Kind = SyntaxKind.ImplementsKeyword Then
                implementsClause = ParseImplementsList()

                If kind = SyntaxKind.DelegateSubStatement Then
                    ' davidsch - This error was reported in Declared in Dev10
                    implementsClause = ReportSyntaxError(implementsClause, ERRID.ERR_DelegateCantImplement)
                End If
            End If
        End Sub


    End Class

End Namespace

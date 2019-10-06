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
        ' *     Parser::ParseFunctionDeclaration
        ' *
        ' * Purpose:
        ' *     Parses a Function definition.
        ' *
        ' **********************************************************************/

        ' [in] specifiers on definition
        ' [in] token starting definition
        ' File: Parser.cpp
        ' Lines: 8470 - 8470
        ' MethodDeclarationStatement* .Parser::ParseFunctionDeclaration( [ ParseTree::AttributeSpecifierList* Attributes ] [ ParseTree::SpecifierList* Specifiers ] [ _In_ Token* Start ] [ bool IsDelegate ] [ _Inout_ bool& ErrorInConstruct ] )

        Private Function ParseFunctionStatement(
                attributes As CoreInternalSyntax.SyntaxList(Of AttributeListSyntax),
                modifiers As CoreInternalSyntax.SyntaxList(Of KeywordSyntax)
            ) As MethodStatementSyntax

            Dim functionKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)

            Debug.Assert(functionKeyword.Kind = SyntaxKind.FunctionKeyword, "Function parsing lost.")

            GetNextToken()

            Dim save_isInMethodDeclarationHeader As Boolean = _isInMethodDeclarationHeader
            _isInMethodDeclarationHeader = True

            Dim save_isInAsyncMethodDeclarationHeader As Boolean = _isInAsyncMethodDeclarationHeader
            Dim save_isInIteratorMethodDeclarationHeader As Boolean = _isInIteratorMethodDeclarationHeader

            _isInAsyncMethodDeclarationHeader = modifiers.Any(SyntaxKind.AsyncKeyword)
            _isInIteratorMethodDeclarationHeader = modifiers.Any(SyntaxKind.IteratorKeyword)

            ' Dev10_504604 we are parsing a method declaration and will need to let the scanner know
            ' that we are so the scanner can correctly identify attributes vs. xml while scanning
            ' the declaration.

            'davidsch - It is not longer necessary to force the scanner state here.  The scanner will
            'only scan xml when the parser explicitly tells it to scan xml.

            Dim name As IdentifierTokenSyntax = Nothing
            Dim genericParams As TypeParameterListSyntax = Nothing
            Dim parameters As ParameterListSyntax = Nothing
            Dim asClause As SimpleAsClauseSyntax = Nothing
            Dim handlesClause As HandlesClauseSyntax = Nothing
            Dim implementsClause As ImplementsClauseSyntax = Nothing

            ParseFunctionOrDelegateStatement(SyntaxKind.FunctionStatement, name, genericParams, parameters, asClause, handlesClause, implementsClause)

            _isInMethodDeclarationHeader = save_isInMethodDeclarationHeader
            _isInAsyncMethodDeclarationHeader = save_isInAsyncMethodDeclarationHeader
            _isInIteratorMethodDeclarationHeader = save_isInIteratorMethodDeclarationHeader

            'Create the Sub statement.
            Dim methodStatement = SyntaxFactory.FunctionStatement(attributes, modifiers, functionKeyword, name, genericParams, parameters, asClause, handlesClause, implementsClause)

            Return methodStatement

        End Function

        Private Sub ParseFunctionOrDelegateStatement(kind As SyntaxKind,
                                                       ByRef ident As IdentifierTokenSyntax,
                                                       ByRef optionalGenericParams As TypeParameterListSyntax,
                                                       ByRef optionalParameters As ParameterListSyntax,
                                                       ByRef asClause As SimpleAsClauseSyntax,
                                                       ByRef handlesClause As HandlesClauseSyntax,
                                                       ByRef implementsClause As ImplementsClauseSyntax)

            Debug.Assert(
                kind = SyntaxKind.FunctionStatement OrElse
                kind = SyntaxKind.DelegateFunctionStatement, "Wrong kind passed to ParseFunctionOrDelegateStatement")

            'TODO - davidsch Can ParseFunctionOrDelegateDeclaration and
            'ParseSubOrDelegateDeclaration share more code? They are nearly the same.

            ' The current token is on the function or delegate's name

            If CurrentToken.Kind = SyntaxKind.NewKeyword Then
                ' "New" gets special attention because attempting to declare a constructor as a
                ' function is, we expect, a common error.
                ident = ParseIdentifierAllowingKeyword()

                ident = ReportSyntaxError(ident, ERRID.ERR_ConstructorFunction)
            Else
                ident = ParseIdentifier()

                ' TODO - davidsch - Why do ParseFunctionDeclaration and ParseSubDeclaration have
                ' different error recovery here?
                If ident.ContainsDiagnostics Then
                    ident = ident.AddTrailingSyntax(ResyncAt({SyntaxKind.OpenParenToken, SyntaxKind.AsKeyword}))
                End If
            End If

            If BeginsGeneric() Then
                optionalGenericParams = ParseGenericParameters()
            End If

            optionalParameters = ParseParameterList()

            ' Check the return type.

            If TryParseSimpleAsClause(asClause, checkForTypeAttributes:=True) Then

            End If

            ' See if we have the HANDLES or the IMPLEMENTS clause on this procedure.

            If CurrentToken.Kind = SyntaxKind.HandlesKeyword Then
                handlesClause = ParseHandlesList()

                If kind = SyntaxKind.DelegateFunctionStatement Then
                    ' davidsch - This error was reported in Declared in Dev10
                    handlesClause = ReportSyntaxError(handlesClause, ERRID.ERR_DelegateCantHandleEvents)
                End If

            ElseIf CurrentToken.Kind = SyntaxKind.ImplementsKeyword Then
                implementsClause = ParseImplementsList()

                If kind = SyntaxKind.DelegateFunctionStatement Then
                    ' davidsch - This error was reported in Declared in Dev10
                    implementsClause = ReportSyntaxError(implementsClause, ERRID.ERR_DelegateCantImplement)
                End If

            End If

        End Sub


    End Class

End Namespace

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
        ' *     Parser::ParseImplementsList
        ' *
        ' **********************************************************************/
        '
        ' File: Parser.cpp
        ' Lines: 8018 - 8018
        ' NameList* .Parser::ParseImplementsList( [ _Inout_ bool& ErrorInConstruct ] )
        '
        Private Function ParseImplementsList() As ImplementsClauseSyntax

            Debug.Assert(CurrentToken.Kind = SyntaxKind.ImplementsKeyword, CalledOnWrongToken(NameOf(ParseImplementsList)))

            Dim implementsKeyword As KeywordSyntax = DirectCast(CurrentToken, KeywordSyntax)
            Dim ImplementsClauses As SeparatedSyntaxListBuilder(Of QualifiedNameSyntax) =
                Me._pool.AllocateSeparated(Of QualifiedNameSyntax)()

            Dim comma As PunctuationSyntax

            GetNextToken()

            Do

                'TODO - davidsch
                ' The old parser did not make a distinction between TypeNames and Names
                ' While there is a ParseTypeName function, the old parser called ParseName.  For now
                ' call ParseName and then break up the name to make a ImplementsClauseItem. The
                ' parameters passed to ParseName guarantee that the name is qualified. The first
                ' parameter ensures qualification.  The last parameter ensures that it is not generic.

                ' AllowGlobalNameSpace
                ' Allow generic arguments

                Dim term = DirectCast(ParseName(
                    requireQualification:=True,
                    allowGlobalNameSpace:=True,
                    allowGenericArguments:=True,
                    allowGenericsWithoutOf:=True,
                    nonArrayName:=True,
                    disallowGenericArgumentsOnLastQualifiedName:=True), QualifiedNameSyntax) ' Disallow generic arguments on last qualified name i.e. on the method name

                ImplementsClauses.Add(term)

                comma = Nothing
                If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then
                    Exit Do
                End If

                ImplementsClauses.AddSeparator(comma)
            Loop

            Dim result = ImplementsClauses.ToList
            Me._pool.Free(ImplementsClauses)

            Return SyntaxFactory.ImplementsClause(implementsKeyword, result)
        End Function

    End Class

End Namespace

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

        ' File: Parser.cpp
        ' Lines: 8062 - 8062
        ' NameList* .Parser::ParseHandlesList( [ _Inout_ bool& ErrorInConstruct ] )
        Private Function ParseHandlesList() As HandlesClauseSyntax

            Debug.Assert(CurrentToken.Kind = SyntaxKind.HandlesKeyword, CalledOnWrongToken(NameOf(ParseHandlesList)))

            Dim handlesKeyword = DirectCast(CurrentToken, KeywordSyntax)
            Dim handlesClauseItems As SeparatedSyntaxListBuilder(Of HandlesClauseItemSyntax) = Me._pool.AllocateSeparated(Of HandlesClauseItemSyntax)()
            Dim comma As PunctuationSyntax

            GetNextToken() ' get off the handles / comma token
            Do
                Dim eventContainer As EventContainerSyntax
                Dim eventMember As IdentifierNameSyntax

                Select Case CurrentToken.Kind
                    Case SyntaxKind.MyBaseKeyword,
                         SyntaxKind.MyClassKeyword,
                         SyntaxKind.MeKeyword

                        eventContainer = SyntaxFactory.KeywordEventContainer(DirectCast(CurrentToken, KeywordSyntax))
                        GetNextToken()

                    Case SyntaxKind.GlobalKeyword
                        ' A handles name can't start with Global, it is local.
                        ' Produce the error, ignore the token and let the name parse for sync.

                        ' we are not consuming Global keyword here as the only acceptable keywords are: Me, MyBase, MyClass
                        eventContainer = SyntaxFactory.WithEventsEventContainer(InternalSyntaxFactory.MissingIdentifier())
                        eventContainer = ReportSyntaxError(eventContainer, ERRID.ERR_NoGlobalInHandles)

                    Case Else
                        eventContainer = SyntaxFactory.WithEventsEventContainer(ParseIdentifier())

                End Select

                Dim Dot As PunctuationSyntax = Nothing

                ' allow implicit line continuation after '.' in handles list - dev10_503311
                If TryGetTokenAndEatNewLine(SyntaxKind.DotToken, Dot, createIfMissing:=True) Then
                    eventMember = InternalSyntaxFactory.IdentifierName(ParseIdentifierAllowingKeyword())

                    ' check if we actually have "withEventsMember.Property.Event"
                    Dim identContainer = TryCast(eventContainer, WithEventsEventContainerSyntax)
                    Dim secondDot As PunctuationSyntax = Nothing

                    If identContainer IsNot Nothing AndAlso TryGetTokenAndEatNewLine(SyntaxKind.DotToken, secondDot, createIfMissing:=True) Then
                        ' former member and dot are shifted into property container.
                        eventContainer = SyntaxFactory.WithEventsPropertyEventContainer(identContainer, Dot, eventMember)
                        ' secondDot becomes the event's dot
                        Dot = secondDot
                        ' parse another event member since the former one has become a property
                        eventMember = InternalSyntaxFactory.IdentifierName(ParseIdentifierAllowingKeyword())
                    End If

                Else
                    eventMember = InternalSyntaxFactory.IdentifierName(InternalSyntaxFactory.MissingIdentifier())
                End If

                Dim item As HandlesClauseItemSyntax = SyntaxFactory.HandlesClauseItem(eventContainer, Dot, eventMember)

                If eventContainer.ContainsDiagnostics OrElse Dot.ContainsDiagnostics OrElse eventMember.ContainsDiagnostics Then

                    If CurrentToken.Kind <> SyntaxKind.CommaToken Then
                        item = ResyncAt(item, SyntaxKind.CommaToken)
                    End If
                End If

                handlesClauseItems.Add(item)

                comma = Nothing
                If Not TryGetTokenAndEatNewLine(SyntaxKind.CommaToken, comma) Then
                    Exit Do
                End If

                handlesClauseItems.AddSeparator(comma)
            Loop

            Dim result = handlesClauseItems.ToList
            Me._pool.Free(handlesClauseItems)

            Return SyntaxFactory.HandlesClause(handlesKeyword, result)
        End Function

    End Class

End Namespace

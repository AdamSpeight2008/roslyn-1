 ' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports System.Text
Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFacts
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Scanner

#Region "EmbeddedToken"

        Private Function XmlMakeBeginEmbeddedToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Const _BeginEm_ = "<%="
            Return Make_XmlToken(_BeginEm_, SyntaxKind.LessThanPercentEqualsToken, precedingTrivia)
        End Function

        Private Function XmlMakeEndEmbeddedToken(
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                  scanTrailingTrivia As ScanTriviaFunc
                                                ) As PunctuationSyntax
            Debug.Assert(Peek().IsEither("%"c, FULLWIDTH_PERCENT_SIGN))
            Debug.Assert(Peek(1) = ">"c)

            Dim spelling As String
            If Peek() = "%"c Then
                AdvanceChar(2)
                spelling = "%>"
            Else
                spelling = GetText(2)
            End If

            Dim followingTrivia = scanTrailingTrivia()
            Return MakePunctuationToken(SyntaxKind.PercentGreaterThanToken, spelling, precedingTrivia, followingTrivia)
        End Function

#End Region

    End Class

End Namespace

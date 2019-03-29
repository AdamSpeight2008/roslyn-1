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

#Region "ProcessingInstruction"

        Private Function XmlMakeBeginProcessingInstructionToken(
                                                                 precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                                 scanTrailingTrivia As ScanTriviaFunc
                                                               ) As PunctuationSyntax
            Const _BeginPI_ = "<?"
            Return Make_XmlToken(_BeginPI_, SyntaxKind.LessThanQuestionToken, precedingTrivia, scanTrailingTrivia)
        End Function

        Private Function XmlMakeProcessingInstructionToken(
                                                            precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                            TokenWidth As Integer
                                                          ) As XmlTextTokenSyntax
            Debug.Assert(TokenWidth > 0)
            'TODO - Normalize new lines.
            Dim text = GetTextNotInterned(TokenWidth)
            Return SyntaxFactory.XmlTextLiteralToken(text, text, precedingTrivia.Node, Nothing)
        End Function

        Private Function XmlMakeEndProcessingInstructionToken(
                                                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                                             ) As PunctuationSyntax
            Const _EndPI_ = "?>"
            Return Make_XmlToken(_EndPi_, SyntaxKind.QuestionGreaterThanToken, precedingTrivia)
        End Function

#End Region

    End Class

End Namespace

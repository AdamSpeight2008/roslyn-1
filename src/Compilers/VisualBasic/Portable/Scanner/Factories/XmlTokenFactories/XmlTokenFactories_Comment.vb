 ' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Scanner

#Region "Comment"
        
        Private Function XmlMakeBeginCommentToken(
                                                   precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                   scanTrailingTrivia As ScanTriviaFunc
                                                 ) As PunctuationSyntax
            Const _BeginComment_ = "<!--"
            Return Make_XmlToken(_BeginComment_, SyntaxKind.LessThanExclamationMinusMinusToken, precedingTrivia, scanTrailingTrivia)
        End Function

        Private Function XmlMakeCommentToken(
                                              precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                              tokenWidth As Integer
                                            ) As XmlTextTokenSyntax
            Debug.Assert(tokenWidth > 0)
            ' TODO - Normalize new lines.
            Dim text = GetTextNotInterned(tokenWidth)
            Return SyntaxFactory.XmlTextLiteralToken(text, text, precedingTrivia.Node, Nothing)
        End Function

        Private Function XmlMakeEndCommentToken( precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode) ) As PunctuationSyntax
            Const _EndComment_ = "-->"
            Return Make_XmlToken(_EndComment_, SyntaxKind.MinusMinusGreaterThanToken, precedingTrivia)
         End Function

#End Region

    End Class

End Namespace

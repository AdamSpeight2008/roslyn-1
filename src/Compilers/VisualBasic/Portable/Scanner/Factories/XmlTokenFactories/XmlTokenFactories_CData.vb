 ' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports System.Text
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Scanner

#Region "CData"
        Private Function XmlMakeBeginCDataToken(
                                                 precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                                 scanTrailingTrivia As ScanTriviaFunc
                                               ) As PunctuationSyntax
            Const _CDATA_ = "<![CDATA["
            Return Make_XmlToken(_CDATA_, SyntaxKind.BeginCDataToken, precedingTrivia, scanTrailingTrivia)
        End Function

        Private Function XmlMakeCDataToken(
                                            precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode),
                                            TokenWidth As Integer,
                                            scratch As StringBuilder
                                          ) As XmlTextTokenSyntax
            Return XmlMakeTextLiteralToken(precedingTrivia, TokenWidth, scratch)
        End Function

        Private Function XmlMakeEndCDataToken(
                                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                             ) As PunctuationSyntax
            Const _EndOfCDATA_ = "]]>"
            Return Make_XmlToken(_EndOfCDATA_, SyntaxKind.EndCDataToken, precedingTrivia)
        End Function
#End Region

    End Class

End Namespace

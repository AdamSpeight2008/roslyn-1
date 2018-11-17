 ' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'-----------------------------------------------------------------------------
' Contains the definition of the Scanner, which produces tokens from text 
'-----------------------------------------------------------------------------
Option Compare Binary

Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax

    Partial Friend Class Scanner

#Region "DTD"

        Private Function XmlMakeBeginDTDToken(
                                               precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                             ) As BadTokenSyntax
            Debug.Assert(NextAre("<!DOCTYPE"))
            Return XmlMakeBadToken(SyntaxSubKind.BeginDocTypeToken, precedingTrivia, 9, ERRID.ERR_DTDNotSupported)
        End Function

        Private Function XmlLessThanExclamationToken(
                                                      state As ScannerState,
                                                      precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                                    ) As BadTokenSyntax
            Debug.Assert(NextAre("<!"))
            Return XmlMakeBadToken(SyntaxSubKind.LessThanExclamationToken, precedingTrivia, 2, If(state = ScannerState.DocType, ERRID.ERR_DTDNotSupported, ERRID.ERR_Syntax))
        End Function

        Private Function XmlMakeOpenBracketToken(
                                                  state As ScannerState,
                                                  precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                                ) As BadTokenSyntax
            Debug.Assert(Peek() = "["c)
            Return XmlMakeBadToken(SyntaxSubKind.OpenBracketToken, precedingTrivia, 1, If(state = ScannerState.DocType, ERRID.ERR_DTDNotSupported, ERRID.ERR_IllegalXmlNameChar))
        End Function

        Private Function XmlMakeCloseBracketToken(
                                                   state As ScannerState,
                                                   precedingTrivia As CoreInternalSyntax.SyntaxList(Of VisualBasicSyntaxNode)
                                                 ) As BadTokenSyntax
            Debug.Assert(Peek() = "]"c)

            Return XmlMakeBadToken(SyntaxSubKind.CloseBracketToken, precedingTrivia, 1, If(state = ScannerState.DocType, ERRID.ERR_DTDNotSupported, ERRID.ERR_IllegalXmlNameChar))
        End Function

#End Region

    End Class

End Namespace

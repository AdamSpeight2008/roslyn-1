﻿' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Threading
Imports Microsoft.CodeAnalysis.Editor.Implementation.Highlighting
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.Editor.VisualBasic.KeywordHighlighting
    <ExportHighlighter(LanguageNames.VisualBasic)>
    Friend Class XmlElementHighlighter
        Inherits AbstractKeywordHighlighter(Of XmlNodeSyntax)

        Protected Overloads Overrides Iterator Function GetHighlights(node As XmlNodeSyntax, cancellationToken As CancellationToken) As IEnumerable(Of TextSpan)
            If cancellationToken.IsCancellationRequested Then
                Return
            End If
            Dim xmlElement = node.GetAncestor(Of XmlElementSyntax)()
            With xmlElement
                If xmlElement IsNot Nothing AndAlso
                   Not .ContainsDiagnostics AndAlso
                   Not .HasAncestor(Of DocumentationCommentTriviaSyntax)() Then

                    With .StartTag
                        If .Attributes.Count = 0 Then
                            Yield .Span
                        Else
                            Yield TextSpan.FromBounds(.LessThanToken.SpanStart, .Name.Span.End)
                            Yield .GreaterThanToken.Span
                        End If
                    End With
                    Yield .EndTag.Span
                End If

            End With
        End Function
    End Class
End Namespace

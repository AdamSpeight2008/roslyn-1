﻿' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Threading
Imports Microsoft.CodeAnalysis.Editor.Implementation.Highlighting
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.Editor.VisualBasic.KeywordHighlighting
    <ExportHighlighter(LanguageNames.VisualBasic)>
    Friend Class ConstructorDeclarationHighlighter
        Inherits AbstractKeywordHighlighter(Of SyntaxNode)

        Protected Overloads Overrides Iterator Function GetHighlights(node As SyntaxNode, cancellationToken As CancellationToken) As IEnumerable(Of TextSpan)
            If cancellationToken.IsCancellationRequested Then
                Return
            End If
            Dim methodBlock = node.GetAncestor(Of MethodBlockBaseSyntax)()
            If methodBlock Is Nothing OrElse Not TypeOf methodBlock.BlockStatement Is SubNewStatementSyntax Then
                Return
            End If

            With methodBlock
                With DirectCast(.BlockStatement, SubNewStatementSyntax)
                    Dim firstKeyword = If(.Modifiers.Count > 0, .Modifiers.First(), .DeclarationKeyword)
                    Yield TextSpan.FromBounds(firstKeyword.SpanStart, .NewKeyword.Span.End)
                End With

                For Each highlight In .GetRelatedStatementHighlights(blockKind:=SyntaxKind.SubKeyword, checkReturns:=True)
                    If cancellationToken.IsCancellationRequested Then
                        Return
                    End If
                    Yield highlight
                Next

                Yield .EndBlockStatement.Span

            End With
        End Function
    End Class
End Namespace

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.Editor.VisualBasic.LineCommit
    Partial Friend Class ContainingStatementInfo
        Private Class MatchingStatementsVisitor
            Inherits VisualBasicSyntaxVisitor(Of IEnumerable(Of StatementSyntax))

            Public Shared ReadOnly Instance As MatchingStatementsVisitor = New MatchingStatementsVisitor()

            Private Sub New()
            End Sub

            Public Overrides Iterator Function VisitClassBlock(node As ClassBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitMethodBlock(node As MethodBlockSyntax) As IEnumerable(Of StatementSyntax)
                yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitConstructorBlock(node As ConstructorBlockSyntax) As IEnumerable(Of StatementSyntax)
               Yield node.BlockStatement
               Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitOperatorBlock(node As OperatorBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitAccessorBlock(node As AccessorBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitDoLoopBlock(node As DoLoopBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.DoStatement
                Yield node.LoopStatement
            End Function

            Public Overrides Iterator Function VisitEnumBlock(node As EnumBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.EnumStatement
                Yield node.EndEnumStatement
            End Function

            Public Overrides Iterator Function VisitForBlock(node As ForBlockSyntax) As IEnumerable(Of StatementSyntax)
                ' TODO: evilness around ending multiple statements at once with a single "next"
                Yield node.ForStatement
                Yield node.NextStatement
            End Function

            Public Overrides Iterator Function VisitForEachBlock(node As ForEachBlockSyntax) As IEnumerable(Of StatementSyntax)
                ' TODO: evilness around ending multiple statements at once with a single "next"
                Yield node.ForEachStatement
                Yield node.NextStatement
            End Function

            Public Overrides Iterator Function VisitMultiLineIfBlock(node As MultiLineIfBlockSyntax) As IEnumerable(Of StatementSyntax)

                Yield node.IfStatement
                Yield node.EndIfStatement
                For Each n in node.ElseIfBlocks.Select(Function(elseIfBlock) elseIfBlock.ElseIfStatement)
                    Yield n
                Next

                If node.ElseBlock IsNot Nothing Then Yield node.ElseBlock.ElseStatement
            End Function

            Public Overrides Iterator Function VisitInterfaceBlock(node As InterfaceBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitModuleBlock(node As ModuleBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitNamespaceBlock(node As NamespaceBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.NamespaceStatement
                Yield node.EndNamespaceStatement
            End Function

            Public Overrides Iterator Function VisitPropertyBlock(node As PropertyBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.PropertyStatement
                Yield node.EndPropertyStatement
            End Function

            Public Overrides Iterator Function VisitSelectBlock(node As SelectBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.SelectStatement
                Yield node.EndSelectStatement
            End Function

            Public Overrides Iterator Function VisitSyncLockBlock(node As SyncLockBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.SyncLockStatement
                Yield node.EndSyncLockStatement
            End Function

            Public Overrides Iterator Function VisitTryBlock(node As TryBlockSyntax) As IEnumerable(Of StatementSyntax)
                yield node.TryStatement
                Yield node.EndTryStatement
                for Each n in node.CatchBlocks.Select(Function(catchBlock) catchBlock.CatchStatement)
                    yield n
                Next
                If node.FinallyBlock IsNot Nothing Then yield node.FinallyBlock.FinallyStatement
            End Function

            Public Overrides Iterator Function VisitStructureBlock(node As StructureBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.BlockStatement
                Yield node.EndBlockStatement
            End Function

            Public Overrides Iterator Function VisitUsingBlock(node As UsingBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.UsingStatement
                Yield node.EndUsingStatement
            End Function

            Public Overrides Iterator Function VisitWithBlock(node As WithBlockSyntax) As IEnumerable(Of StatementSyntax)
                Yield node.WithStatement
                Yield node.EndWithStatement
            End Function
        End Class
    End Class
End Namespace

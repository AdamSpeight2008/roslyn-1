' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Generic
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic
    ''' <summary>
    ''' A region analysis walker that records declared variables.
    ''' </summary>
    Friend Class VariablesDeclaredWalker
        Inherits AbstractRegionControlFlowPass
        Implements IDisposable

        Friend Overloads Shared Function Analyze(info As FlowAnalysisInfo, region As FlowAnalysisRegionInfo) As IEnumerable(Of Symbol)
            Dim walker = New VariablesDeclaredWalker(info, region)
            Try
                Return If(walker.Analyze(), walker._variablesDeclared, SpecializedCollections.EmptyEnumerable(Of Symbol)())
            Finally
                walker.Free()
            End Try
        End Function

        Private ReadOnly _variablesDeclared As PooledHashSet(Of Symbol) = s_pool.Allocate()

        Private Overloads Function Analyze() As Boolean
            ' only one pass needed.
            Return Scan()
        End Function

        Private Sub New(info As FlowAnalysisInfo, region As FlowAnalysisRegionInfo)
            MyBase.New(info, region)
        End Sub

        Public Overrides Function VisitLocalDeclaration(node As BoundLocalDeclaration) As BoundNode
            If IsInside Then
                _variablesDeclared.Add(node.LocalSymbol)
            End If
            Return MyBase.VisitLocalDeclaration(node)
        End Function

        Protected Overrides Sub VisitForStatementVariableDeclaration(node As BoundForStatement)
            If IsInside AndAlso
                    node.DeclaredOrInferredLocalOpt IsNot Nothing Then
                _variablesDeclared.Add(node.DeclaredOrInferredLocalOpt)
            End If
            MyBase.VisitForStatementVariableDeclaration(node)
        End Sub

        Public Overrides Function VisitLambda(node As BoundLambda) As BoundNode
            If IsInside Then
                For Each parameter In node.LambdaSymbol.Parameters
                    _variablesDeclared.Add(parameter)
                Next
            End If

            Return MyBase.VisitLambda(node)
        End Function

        Public Overrides Function VisitQueryableSource(node As BoundQueryableSource) As BoundNode
            MyBase.VisitQueryableSource(node)

            If Not node.WasCompilerGenerated AndAlso node.RangeVariables.Length > 0 AndAlso IsInside Then
                Debug.Assert(node.RangeVariables.Length = 1)
                _variablesDeclared.Add(node.RangeVariables(0))
            End If

            Return Nothing
        End Function

        Public Overrides Function VisitRangeVariableAssignment(node As BoundRangeVariableAssignment) As BoundNode
            If Not node.WasCompilerGenerated AndAlso IsInside Then
                _variablesDeclared.Add(node.RangeVariable)
            End If

            MyBase.VisitRangeVariableAssignment(node)
            Return Nothing
        End Function

        Protected Overrides Sub VisitCatchBlock(catchBlock As BoundCatchBlock, ByRef finallyState As LocalState)
            If IsInsideRegion(catchBlock.Syntax.Span) Then
                If catchBlock.LocalOpt IsNot Nothing Then
                    _variablesDeclared.Add(catchBlock.LocalOpt)
                End If

            End If

            MyBase.VisitCatchBlock(catchBlock, finallyState)
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    _variablesDeclared.Free()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace

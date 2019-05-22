' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Generic
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.PooledObjects

Namespace Microsoft.CodeAnalysis.VisualBasic

    ''' <summary>
    ''' An analysis that computes the set of variables that may be used
    ''' before being assigned anywhere within a method.
    ''' </summary>
    ''' <remarks></remarks>
    Friend NotInheritable Class UnassignedVariablesWalker
        Inherits DataFlowPass
        Implements IDisposable

        ' TODO: normalize the result by removing variables that are unassigned in an unmodified flow analysis.
        Private Sub New(info As FlowAnalysisInfo)
            MyBase.New(info, suppressConstExpressionsSupport:=False, trackStructsWithIntrinsicTypedFields:=True)
        End Sub

        Friend Overloads Shared Function Analyze(info As FlowAnalysisInfo) As PooledHashSet(Of Symbol)
            Dim walker = New UnassignedVariablesWalker(info)
            Try
                Return If(walker.Analyze(), walker._result, s_pool.Allocate())
            Finally
                walker.Free()
            End Try
        End Function

        Private ReadOnly _result As PooledHashSet(Of Symbol) = s_pool.Allocate()

        Protected Overrides Sub ReportUnassigned(local As Symbol,
                                                 node As SyntaxNode,
                                                 rwContext As ReadWriteContext,
                                                 Optional slot As Integer = SlotKind.NotTracked,
                                                 Optional boundFieldAccess As BoundFieldAccess = Nothing)

            Debug.Assert(local.Kind <> SymbolKind.Field OrElse boundFieldAccess IsNot Nothing)

            If local.Kind = SymbolKind.Field Then
                Dim sym As Symbol = GetNodeSymbol(boundFieldAccess)

                ' Ambiguous implicit receiver with should not even considered to be unassigned
                Debug.Assert(Not TypeOf sym Is AmbiguousLocalsPseudoSymbol)

                If sym IsNot Nothing Then
                    _result.Add(sym)
                End If

            Else
                _result.Add(local)
            End If

            MyBase.ReportUnassigned(local, node, rwContext, slot, boundFieldAccess)
        End Sub

        Protected Overrides ReadOnly Property SuppressRedimOperandRvalueOnPreserve As Boolean
            Get
                Return False
            End Get
        End Property

        Friend Overrides Sub AssignLocalOnDeclaration(local As LocalSymbol, node As BoundLocalDeclaration)
            ' NOTE: static locals should not be considered assigned even in presence of initializer
            If Not local.IsStatic Then
                MyBase.AssignLocalOnDeclaration(local, node)
            End If
        End Sub

        Protected Overrides ReadOnly Property IgnoreOutSemantics As Boolean
            Get
                Return False
            End Get
        End Property

        Protected Overrides ReadOnly Property EnableBreakingFlowAnalysisFeatures As Boolean
            Get
                Return True
            End Get
        End Property

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    _result.Free()
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

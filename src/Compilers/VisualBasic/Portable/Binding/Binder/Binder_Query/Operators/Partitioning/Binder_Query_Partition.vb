' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' This helper method does all the work to bind Take and Skip query operators.
        ''' </summary>
        Private Function BindPartitionClause(
                                              source As BoundQueryClauseBase,
                                              partition As PartitionClauseSyntax,
                                              operatorName As String,
                                              diagnostics As DiagnosticBag
                                            ) As BoundQueryClause
            Debug.Assert(source IsNot Nothing, NameOf(source))
            Debug.Assert(partition IsNot Nothing, NameOf(partition))

            ' Bind the Count expression as a value, conversion should take care of the rest (making it an RValue, etc.). 
            Dim boundCount As BoundExpression = Me.BindValue(partition.Count, diagnostics)

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If source.Type.IsErrorType() Then
                boundCallOrBadExpression = BadExpression(partition, ImmutableArray.Create(source, boundCount),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If boundCount.HasErrors OrElse boundCount.Type?.IsErrorType() Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If

                boundCallOrBadExpression = BindQueryOperatorCall(partition, source,
                                                               operatorName,
                                                               ImmutableArray.Create(boundCount),
                                                               partition.SkipOrTakeKeyword.Span,
                                                               diagnostics)

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            End If

            Return New BoundQueryClause(partition,
                                        boundCallOrBadExpression,
                                        source.RangeVariables,
                                        source.CompoundVariableType,
                                        ImmutableArray(Of Binder).Empty,
                                        boundCallOrBadExpression.Type)
        End Function

    End Class

End Namespace

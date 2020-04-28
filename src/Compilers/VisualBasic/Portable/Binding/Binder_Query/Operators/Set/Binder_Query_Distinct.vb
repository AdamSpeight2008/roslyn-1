' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Distinct operator.
        ''' 
        '''     {Preceding query operators} Distinct
        ''' 
        ''' Ex: From a In AA Distinct ==> AA.Distinct()
        ''' 
        ''' </summary>
        Private Function BindDistinctClause(
                                             source As BoundQueryClauseBase,
                                             distinct As DistinctClauseSyntax,
                                             diagnostics As DiagnosticBag
                                           ) As BoundQueryClause

            Dim boundCallOrBadExpression As BoundExpression

            If source.Type.IsErrorType() Then
                ' Operator BindQueryClauseCall will fail, let's not bother.
                boundCallOrBadExpression = BadExpression(distinct,
                                                         source,
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                boundCallOrBadExpression = BindQueryOperatorCall(distinct,
                                                                 source,
                                                                 StringConstants.DistinctMethod,
                                                                 ImmutableArray(Of BoundExpression).Empty,
                                                                 distinct.DistinctKeyword.Span,
                                                                 diagnostics)
            End If

            Return New BoundQueryClause(distinct,
                                        boundCallOrBadExpression,
                                        source.RangeVariables,
                                        source.CompoundVariableType,
                                        ImmutableArray(Of Binder).Empty,
                                        boundCallOrBadExpression.Type)
        End Function

    End Class

End Namespace

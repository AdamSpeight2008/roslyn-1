' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Where operator.
        ''' 
        '''     {Preceding query operators} Where {expression}
        ''' 
        ''' Ex: From a In AA Where a > 0 ==> AA.Where(Function(a) a > b)
        ''' 
        ''' </summary>
        Private Function BindWhereClause(
                                          source As BoundQueryClauseBase,
                                          where As WhereClauseSyntax,
                                          diagnostics As DiagnosticBag
                                        ) As BoundQueryClause
            Return BindFilterQueryOperator(source,
                                           where,
                                           StringConstants.WhereMethod,
                                           where.WhereKeyword.Span,
                                           where.Condition,
                                           diagnostics)
        End Function

    End Class

End Namespace

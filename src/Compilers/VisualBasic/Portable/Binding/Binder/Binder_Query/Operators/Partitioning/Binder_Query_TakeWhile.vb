' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Take While operator.
        ''' 
        '''     {Preceding query operators} Take While {expression}
        ''' 
        ''' Ex: From a In AA Skip While a > 0 ==> AA.TakeWhile(Function(a) a > b)
        ''' 
        ''' </summary>
        Private Function BindTakeWhileClause(
                                              source As BoundQueryClauseBase,
                                              takeWhile As PartitionWhileClauseSyntax,
                                              diagnostics As DiagnosticBag
                                            ) As BoundQueryClause
            Return BindFilterQueryOperator(source,
                                           takeWhile,
                                           StringConstants.TakeWhileMethod,
                                           GetQueryOperatorNameSpan(takeWhile.SkipOrTakeKeyword, takeWhile.WhileKeyword),
                                           takeWhile.Condition,
                                           diagnostics)
        End Function

    End Class

End Namespace

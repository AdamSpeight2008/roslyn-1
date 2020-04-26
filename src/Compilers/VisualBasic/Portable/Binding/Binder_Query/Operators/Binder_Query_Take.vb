' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Take operator.
        ''' 
        '''     {Preceding query operators} Take {expression}
        ''' 
        ''' Ex: From a In AA Take 10 ==> AA.Take(10)
        ''' 
        ''' </summary>
        Private Function BindTakeClause(
                                         source As BoundQueryClauseBase,
                                         take As PartitionClauseSyntax,
                                         diagnostics As DiagnosticBag
                                       ) As BoundQueryClause
            Return BindPartitionClause(source, take, StringConstants.TakeMethod, diagnostics)
        End Function

    End Class

End Namespace

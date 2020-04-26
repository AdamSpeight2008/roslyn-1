' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given result of binding preceding query operators, the source, bind the following Skip operator.
        ''' 
        '''     {Preceding query operators} Skip {expression}
        ''' 
        ''' Ex: From a In AA Skip 10 ==> AA.Skip(10)
        ''' 
        ''' </summary>
        Private Function BindSkipClause(
                                         source As BoundQueryClauseBase,
                                         skip As PartitionClauseSyntax,
                                         diagnostics As DiagnosticBag
                                       ) As BoundQueryClause
            Return BindPartitionClause(source, skip, StringConstants.SkipMethod, diagnostics)
        End Function

    End Class

End Namespace

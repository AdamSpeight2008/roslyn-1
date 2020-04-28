' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        Private Function BindZipClause(
                                        source As BoundQueryClauseBase,
                                        zipClause As ZipClauseSyntax,
                                        operators As SyntaxList(Of QueryClauseSyntax).Enumerator,
                                        diagnostics As DiagnosticBag
                                      ) As BoundQueryClauseBase
            ' 
            ' source0 ZIP source1 [otherclauses]
            '
            Dim zipWithNode As VisualBasicSyntaxNode = LambdaUtilities.GetZipLambdaBody(zipClause)


            '' Create LambdaSymbol for the shape of the many-selector lambda.
            '' Create LambdaSymbol for the shape of the selector.
            Dim zipWithParam As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                       source.CompoundVariableType,
                                                                                       zipWithNode, source.RangeVariables)
            Dim lambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.ZipQueryLambda,zipWithNode),
                                                        ImmutableArray.Create(zipWithParam))
            Dim zipRange = BindCollectionRangeVariable(zipClause.Variables(0),False,nothing, Diagnostics)
            ' Create binder for the selector.
            Dim zipBinder As New QueryLambdaBinder(lambdaSymbol,  source.RangeVariables)
            Dim zip As BoundExpression = zipBinder.BindZipClauseSelector(source, zipClause, operators, diagnostics)


            Dim zipLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                   source.RangeVariables,
                                                   zip,
                                                   exprIsOperandOfConditionalBranch:=False)

            lambdaSymbol.SetQueryLambdaReturnType(zip.Type)
            zipLambda.SetWasCompilerGenerated()

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If source.Type.IsErrorType() Then
                boundCallOrBadExpression = BadExpression(zipClause, ImmutableArray.Create(Of BoundExpression)(source, zipLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                Dim suppressDiagnostics As DiagnosticBag = Nothing

                If ShouldSuppressDiagnostics(zipLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If

                boundCallOrBadExpression = BindQueryOperatorCall(zipClause, source,
                                                                 StringConstants.ZipMethod,
                                                                 ImmutableArray.Create(Of BoundExpression)(zipLambda, zipRange),
                                                                 zipClause.ZipKeyword.Span,
                                                                 diagnostics)

                If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            End If

            Return New BoundQueryClause(zipClause,
                                        boundCallOrBadExpression, ImmutableArray.Create(Of RangeVariableSymbol)(),
                                        zipLambda.Expression.Type,
                                        ImmutableArray.Create(Of Binder)(zipBinder),
                                        boundCallOrBadExpression.Type)
        End Function

    End Class

End Namespace

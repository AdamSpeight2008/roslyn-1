' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' This helper method does all the work to bind Where, Take While and Skip While query operators.
        ''' </summary>
        Private Function BindFilterQueryOperator(
                                                  source As BoundQueryClauseBase,
                                                  operatorSyntax As QueryClauseSyntax,
                                                  operatorName As String,
                                                  operatorNameLocation As TextSpan,
                                                  condition As ExpressionSyntax,
                                                  diagnostics As DiagnosticBag
                                                ) As BoundQueryClause

            Dim suppressDiagnostics As DiagnosticBag = Nothing

            ' Create LambdaSymbol for the shape of the filter lambda.

            Dim param As BoundLambdaParameterSymbol = CreateQueryLambdaParameterSymbol(GetQueryLambdaParameterName(source.RangeVariables), 0,
                                                                                       source.CompoundVariableType,
                                                                                       condition, source.RangeVariables)

            Debug.Assert(LambdaUtilities.IsLambdaBody(condition))
            Dim lambdaSymbol = Me.CreateQueryLambdaSymbol((SynthesizedLambdaKind.FilterConditionQueryLambda,condition),
                                                          ImmutableArray.Create(param))

            ' Create binder for a filter condition.
            Dim filterBinder As New QueryLambdaBinder(lambdaSymbol, source.RangeVariables)

            ' Bind condition as a value, conversion should take care of the rest (making it an RValue, etc.). 
            Dim predicate As BoundExpression = filterBinder.BindValue(condition, diagnostics)

            ' Need to verify result type of the condition and enforce ExprIsOperandOfConditionalBranch for possible future conversions. 
            ' In order to do verification, we simply attempt conversion to boolean in the same manner as BindBooleanExpression.
            Dim conversionDiagnostic = DiagnosticBag.GetInstance()

            Dim boolSymbol As NamedTypeSymbol = GetSpecialType(SpecialType.System_Boolean, condition, diagnostics)

            ' If predicate has type Object we will keep result of conversion, otherwise we drop it.
            Dim predicateType As TypeSymbol = predicate.Type
            Dim keepConvertedPredicate As Boolean = False

            If predicateType Is Nothing Then
                If predicate.IsNothingLiteral() Then
                    keepConvertedPredicate = True
                End If
            ElseIf predicateType.IsObjectType() Then
                keepConvertedPredicate = True
            End If

            Dim convertedToBoolean As BoundExpression = filterBinder.ApplyImplicitConversion(condition,
                                                                                             boolSymbol, predicate,
                                                                                             conversionDiagnostic, isOperandOfConditionalBranch:=True)

            ' If we don't keep result of the conversion, keep diagnostic if conversion failed.
            If keepConvertedPredicate Then
                predicate = convertedToBoolean
                diagnostics.AddRange(conversionDiagnostic)

            ElseIf convertedToBoolean.HasErrors AndAlso conversionDiagnostic.HasAnyErrors() Then
                diagnostics.AddRange(conversionDiagnostic)
                ' Suppress any additional diagnostic, otherwise we might end up with duplicate errors.
                If suppressDiagnostics Is Nothing Then
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If
            End If

            conversionDiagnostic.Free()

            ' Bind the Filter
            Dim filterLambda = CreateBoundQueryLambda(lambdaSymbol,
                                                      source.RangeVariables,
                                                      predicate,
                                                      exprIsOperandOfConditionalBranch:=True)

            filterLambda.SetWasCompilerGenerated()

            ' Now bind the call.
            Dim boundCallOrBadExpression As BoundExpression

            If source.Type.IsErrorType() Then
                boundCallOrBadExpression = BadExpression(operatorSyntax, ImmutableArray.Create(Of BoundExpression)(source, filterLambda),
                                                         ErrorTypeSymbol.UnknownResultType).MakeCompilerGenerated()
            Else
                If suppressDiagnostics Is Nothing AndAlso ShouldSuppressDiagnostics(filterLambda) Then
                    ' Operator BindQueryClauseCall will fail, let's suppress any additional errors it will report.
                    suppressDiagnostics = DiagnosticBag.GetInstance()
                    diagnostics = suppressDiagnostics
                End If

                boundCallOrBadExpression = BindQueryOperatorCall(operatorSyntax, source,
                                                                 operatorName,
                                                                 ImmutableArray.Create(Of BoundExpression)(filterLambda),
                                                                 operatorNameLocation,
                                                                 diagnostics)
            End If

            If suppressDiagnostics IsNot Nothing Then suppressDiagnostics.Free()

            Return New BoundQueryClause(operatorSyntax,
                                        boundCallOrBadExpression,
                                        source.RangeVariables,
                                        source.CompoundVariableType,
                                        ImmutableArray.Create(Of Binder)(filterBinder),
                                        boundCallOrBadExpression.Type)
        End Function

    End Class

End Namespace

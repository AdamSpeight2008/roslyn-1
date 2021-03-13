' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.VisualBasic.Emit
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

        Public Overrides Function VisitTypeOf(node As BoundTypeOf) As BoundNode
            Dim f As New SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)
            If node.Local IsNot Nothing Then
                If node.IsTypeOfIsNotExpression Then
                    Return f.BadExpression()
                Else
                    Return f.BadExpression()
                End If
            End If
            Return MyBase.VisitTypeOf(node)
        End Function

        Public Overrides Function VisitTypeOfMany(node As BoundTypeOfMany) As BoundNode

            ' 
            ' A TypeofMany expression is lowered into a sequence of TypeOfExpressions with a boolean logic operations between each pair.
            ' The boolean logic operation for Is is OrElse,
            ' the boolean logic operation for IsNot is AndAlso.
            '
            ' Lowerings:
            '   TypeOf expression Is (Of T0, T1, T2)
            ' => ((TypeOf expression Is T0) OrElse (TypeOf expression Is T1) OrElse (TypeOf expression Is T2))
            '
            '   TypeOe expression IsNot (Of T0, T1, T2)
            ' => (TypeOf expression IsNot T0) AndAlso (TypeOf expression IsNot T1) AndAlso (TypeOf expression IsNot T2))
            '
            Dim f As New SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)
            Dim expr = TryCast(node.Syntax, TypeOfExpressionSyntax)
            If expr Is Nothing Then Return f.BadExpression()

            ' Validate that the operation is valid, either an IsExpression or IsNot Expression.
            Dim is___Operand = expr.OperatorToken.Kind = SyntaxKind.IsExpression
            Dim isNotOperand = expr.OperatorToken.Kind = SyntaxKind.IsNotExpression
            If Not (is___Operand Xor isNotOperand) Then
                ' If both are false or both are true, something has gone wrong.
                _diagnostics.Add(ERRID.ERR_ArgumentSyntax, expr.OperatorToken.GetLocation())
                Return MyBase.VisitTypeOfMany(node)
            End If

            Dim types = TryCast(expr.Type, TypeArgumentListSyntax)
            if types Is Nothing Then Return f.BadExpression()

            Dim numberOfTypes = node.TypeArguments.Arguments.Length
            If numberOfTypes < 1 Then Return MyBase.VisitTypeOfMany(node)

            Return Rewrite_TypeOfMany(node, f, is___Operand, isNotOperand, types, numberOfTypes)
        End Function

        Private Function Rewrite_TypeOfMany( node As BoundTypeOfMany,
                                             f As SyntheticBoundNodeFactory,
                                             isOperand As Boolean,
                                             isNotOperand As Boolean,
                                             types As TypeArgumentListSyntax,
                                             numberOfTypes As Integer
                                           ) As BoundExpression
            ' Prepare to construce the output expression.
            Dim result As BoundExpression = Nothing
            Dim idx = 0
            Dim bool = GetSpecialType(SpecialType.System_Boolean)
            While idx < numberOfTypes
                ' Construct the typeof expression corrisponding to the current iteration's type.
                Dim thisTypeOfExpr = New BoundTypeOf(
                                                      types.Arguments(idx),
                                                      Nothing,
                                                      Nothing,
                                                      node.TypeArguments.Arguments(idx),
                                                      node.Operand,
                                                      isNotOperand,
                                                      bool
                                                    ).MakeRValue.MakeCompilerGenerated
                ' Combine the typeof expression with the previous one.
                If idx = 0 Then
                    ' First typeof expression in list, so isn't combined with anything.
                    result = thisTypeOfExpr
                ElseIf isNotOperand Then
                    ' use an AndAlso for the combining
                    result = f.LogicalAndAlso(result, thisTypeOfExpr)
                ElseIf isOperand Then
                    ' use an OrElse for the combining
                    result = f.LogicalOrElse(result, thisTypeOfExpr)
                Else
                    ' This should be impossible to reach, in case it is thrown an exception.
                    ' So we can investigate further.
                    Throw ExceptionUtilities.Unreachable()
                End If
                idx += 1
            End While

            Return result
        End Function

    End Class

End Namespace

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.VisualBasic.Emit
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

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
            Dim expr = TryCast(node.Syntax, TypeOfManyExpressionSyntax)
            If expr Is Nothing Then
                Return f.BadExpression()
            End If

            ' Validate that the operation is valid, either an IsExpression or IsNot Expression.
            Dim isOperand = expr.OperatorToken.Kind = SyntaxKind.IsExpression
            Dim isNotOperand = expr.OperatorToken.Kind = SyntaxKind.IsNotExpression
            If Not (isOperand Xor isNotOperand) Then
                ' If both are false or both are true, something has gone wrong.
                _diagnostics.Add(ERRID.ERR_ArgumentSyntax, expr.OperatorToken.GetLocation())
                Return MyBase.VisitTypeOfMany(node)
            End If

            Dim numberOfTypes = node.TypeArguments.Arguments.Length
            If numberOfTypes < 1 Then
                Return MyBase.VisitTypeOfMany(node)
            End If
            ' Prepare to construce the output expression.
            Dim result As BoundExpression = Nothing
            Dim idx = 0
            While idx < numberOfTypes
                ' Construct the typeof expression corrisponding to the current iteration's type.
                Dim thisTypeOfExpr = New BoundTypeOf(expr.Types.Arguments(idx), node.TypeArguments.Arguments(idx), node.Operand, isNotOperand, f.[Boolean]).MakeRValue.MakeCompilerGenerated
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

        Public Overrides Function VisitExpressionIntoVariable(node As BoundExpressionIntoVariable) As BoundNode
            ' Currently only TypeOfExpression is supported.
            ' Other expression maybe supported in the future.
            If TypeOf node.Expression Is BoundTypeOf Then
                Return Rewrite_TypeOfExpression_Into_Variable(node, DirectCast(node.Expression, BoundTypeOf))
            End If

            Return MyBase.VisitExpressionIntoVariable(node)
        End Function

        Private Function Rewrite_TypeOfExpression_Into_Variable(node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            Dim targetType = typeofexpr.TargetType
            Dim factory As New SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)

            ' Depending on the type kind of the target type, different lowering are created.
            If targetType.IsReferenceType Then
                Return Rewrite_TypeOf_Into_Variable_With_Class(factory, targetType, node, typeofexpr)
            ElseIf targetType.IsStructureType Then
                Return Rewrite_TypeOf_Into_Variable_With_Structure(factory, targetType, node, typeofexpr)
            End If
            Return MyBase.VisitExpressionIntoVariable(node)
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Class(f As SyntheticBoundNodeFactory,
                                                                 targettype As TypeSymbol,
                                                                 node As BoundExpressionIntoVariable,
                                                                 typeofexpr As BoundTypeOf) As BoundNode
            ' 
            ' The lowering of:=
            '    TypeOf expression Is type Into variable
            ' Should be equivalent to calling the following function.
            '    => Is(Of type)(expression, variable)
            '
            '  Function [Is](Of T As Class)( expression As Object, <Out> ByRef output As T ) As Boolean
            '    Dim output = TryCast(operand, targetType)
            '    Return output. IsNot nothing
            '  End Function
            '

            ' Construct the assignment that passes back the result of the try cast.
            Dim passback = f.AssignmentExpression(node.Variable,
                               f.TryCast(Make_DirectCastToObject(f, node.Syntax, typeofexpr.Operand), targettype))

            ' Construct the result of passback isnot nothing.
            Dim result = f.ReferenceIsNotNothing(passback)

            ' Finally construct the sequence of code.
            Dim code_sequence = f.Sequence(passback, result)
            Return code_sequence
        End Function

        Private Function Make_DirectCastToObject(f As SyntheticBoundNodeFactory, node As SyntaxNode, expression As BoundExpression) As BoundExpression
            ' Construct: DirectCast( expression, Object)
            Return New BoundDirectCast(node, expression, Conversions.ClassifyDirectCastConversion(expression.Type, f.[Object], Nothing), f.[Object])
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Structure(f As SyntheticBoundNodeFactory, targettype As TypeSymbol, node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            ' 
            ' The lowering of:=
            '    TypeOf expression Is type Into variable
            ' Should be equivalent to calling the following function.
            '    => Is(Of type)(expression, variable)
            '
            '   Function Is(Of T As Structure)( expression As Object,<Out> ByRef output As T ) As Boolean ' non-Nullable value type
            '     Dim tmp As T? = DirectCast(expression, T?)
            '     output = tmp.GetValueOrDefault()
            '     Return tmp.HasValue
            '   End Function
            '

            ' Define the type of Nullable(Of T)
            Dim nullableOf_T = f.NullableOf(targettype)

            ' Define a temporary variable of type Nullable(Of T).
            Dim tmp_variable = f.SynthesizedLocal(nullableOf_T)

            ' Construct the DirectCast( expression, Object)
            Dim expression_as_object = Make_DirectCastToObject(f, node.Syntax, typeofexpr.Operand)

            ' Construct the assignment to the temporary variable with DirectCast( expression_as_object, Nullable(Of T))
            Dim assigment = f.AssignmentExpression(f.Local(tmp_variable, True),
                                             New BoundDirectCast(node.Syntax, expression_as_object,
                                             Conversions.ClassifyDirectCastConversion(expression_as_object.Type, nullableOf_T, Nothing), nullableOf_T))

            ' Construct the assignment that passes back the value or the default value from the temporary variable.
            Dim passback = f.AssignmentExpression(node.Variable,
                                                f.Call(f.Local(tmp_variable, False),
                                                       GetNullableMethod(node.Syntax, nullableOf_T, SpecialMember.System_Nullable_T_GetValueOrDefault)))

            ' Construct the result which check that the temporary variable has value.
            Dim result = f.Call(assigment, GetNullableMethod(node.Syntax, nullableOf_T, SpecialMember.System_Nullable_T_get_HasValue))

            ' Finally construct the sequence of code.
            Dim code_sequence = f.Sequence(tmp_variable, assigment, passback, result)
            Return code_sequence
        End Function
    End Class
End Namespace

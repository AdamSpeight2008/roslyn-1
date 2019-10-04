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
            Dim f AS new SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)
            Dim tBoolean = GetSpecialType(SpecialType.System_Boolean)
            Dim expr = TryCast(node.Syntax, TypeOfManyExpressionSyntax)
            Dim isOperand = expr.OperatorToken.Kind = SyntaxKind.IsExpression
            Dim isNotOperand = expr.OperatorToken.Kind = syntaxkind.IsNotExpression
            Dim result As BoundExpression = nothing
            If isOperand = false AndAlso isNotOperand = false Then
                _diagnostics.Add(ERRID.ERR_ArgumentSyntax, expr.OperatorToken.GetLocation())
                return mybase.VisitTypeOfMany(node)
            End If
            Dim idx = 0
            While idx < node.TypeArguments.Arguments.Length
                Dim thisTypeOfExpr = New BoundTypeOf(expr.Types.Arguments(idx), node.TypeArguments.Arguments(idx), node.Operand, isNotOperand, tBoolean).MakeRValue.MakeCompilerGenerated
                If idx = 0 Then
                    result = thisTypeOfExpr
                Else if isnotOperand Then
                    result = f.LogicalAndAlso(result, thisTypeOfExpr)
                else if isOperand Then
                    result = f.LogicalOrElse(result, thisTypeOfExpr)
                End If
                idx += 1
            End While
            Return result
        End Function

        Public Overrides Function VisitExpressionIntoVariable(node As BoundExpressionIntoVariable) As BoundNode
            If TypeOf node.Expression Is BoundTypeOf Then
                Return Rewrite_TypeOfExpression_Into_Variable(node,DirectCast(node.Expression, BoundTypeOf))
            End If
            Return MyBase.VisitExpressionIntoVariable(node)
        End Function

        Private Function Rewrite_TypeOfExpression_Into_Variable(node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            Dim targetType = typeofexpr.TargetType
            Dim factory As New SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)
            If targetType.IsReferenceType Then
                Return Rewrite_TypeOf_Into_Variable_With_Class(factory, targetType, node, typeofexpr)
            ElseIf targetType.IsStructureType Then
                Return Rewrite_TypeOf_Into_Variable_With_Structure(factory, targetType, node, typeofexpr)
            End If
            Return MyBase.VisitExpressionIntoVariable(node)
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Class(f As SyntheticBoundNodeFactory, targettype As TypeSymbol, node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            '  Function [Is](Of T As Class)( expression As Object, <Out> ByRef output As T ) As Boolean
            '    Dim output = TryCast(operand, targetType)
            '    Return output. IsNot nothing
            '  End Function

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
            Dim t_Obj = f.SpecialType(SpecialType.System_Object)
            Dim c_obj = New BoundDirectCast(node, expression, Conversions.ClassifyDirectCastConversion(expression.Type, t_Obj, Nothing), t_Obj)
            Return c_obj
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Structure(f As SyntheticBoundNodeFactory, targettype As TypeSymbol, node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            '  Function Is(Of T As Structure)( expression As Object,<Out> ByRef output As T ) As Boolean ' non-Nullable value type
            '    Dim tmp As T? = DirectCast(expression, T?)
            '    output = tmp.GetValueOrDefault()
            '    Return tmp.HasValue
            '  End Function
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

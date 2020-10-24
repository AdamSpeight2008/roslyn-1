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
            While idx < numberOfTypes
                ' Construct the typeof expression corrisponding to the current iteration's type.
                Dim thisTypeOfExpr = New BoundTypeOf(types.Arguments(idx),
                                                     node.TypeArguments.Arguments(idx),
                                                     node.Operand,
                                                     isNotOperand,
                                                     f.[Boolean]).MakeRValue.MakeCompilerGenerated
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

#Region "TypeOf ... In ..."

        Public Overrides Function VisitExpressionIntoVariable(node As BoundExpressionIntoVariable) As BoundNode
            ' Currently only TypeOfExpression is supported.
            ' Other expression maybe supported in the future.
            If TypeOf node.Expression Is BoundTypeOf Then
                Return Rewrite_TypeOfExpression_Into_Variable(node, DirectCast(node.Expression, BoundTypeOf))
            Else
                Return MyBase.VisitExpressionIntoVariable(node)
            End If
        End Function

        Private Function Can_ConvertTo(f As SyntheticBoundNodeFactory, node as BoundExpression, type As TypeSymbol, byref convkind As ConversionKind) As Boolean
            Dim conversion =ClassifyConversion(f.Compilation, node.Type, type)
            convkind = Conversion.Kind
            Return conversion.Exists
        End Function

        Private Function Rewrite_TypeOfExpression_Into_Variable(node As BoundExpressionIntoVariable, typeofexpr As BoundTypeOf) As BoundNode
            Dim targetType = typeofexpr.TargetType
            Dim factory As New SyntheticBoundNodeFactory(_topMethod, _currentMethodOrLambda, node.Syntax, _compilationState, _diagnostics)
            Dim convKind As ConversionKind = Nothing
            If Not Can_ConvertTo(factory, typeofexpr.Operand, typeofexpr.TargetType, convKind) Then Return factory.BadExpression()

            Dim convkind1 As ConversionKind = nothing
            If Not Can_ConvertTo(factory, typeofexpr.Operand, node.Variable.Type, convKind1) Then Return factory.BadExpression()

            ' Depending on the type kind of the target type, different lowering are created.
            If targetType.IsNullableType Then
                Return factory.BadExpression()
                'Return Rewrite_TypeOf_Into_Variable_With_Nullable(factory, targetType, node, typeofexpr)
            Else If targetType.IsStructureType Then
                Return Rewrite_TypeOf_Into_Variable_With_Structure(factory, targetType, node, typeofexpr)
            Else
                Return Rewrite_TypeOf_Into_Variable_With_Class(factory, targetType, node, typeofexpr)
            End If
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Nullable(factory As SyntheticBoundNodeFactory,
                                                                    targetType As TypeSymbol,
                                                                    intoExpr As BoundExpressionIntoVariable,
                                                                    typeofexpr As BoundTypeOf) As BoundNode
            'Function TypeofIntoNullable(Of T0 As Structure, T1 As Structure)(input As T0?, ByRef output As T1?) As Boolean
            '  output = If(input.HasValue, Microsoft.VisualBasic.CompilerServices.Conversions.ToGenericParameter_T_Object Ctype(Ctype(input.value, Object), T1) ,Nothing)
            '  Return output.HasValue
            'End Function
            Dim input = VisitExpression(typeofexpr.Operand)
            Dim output = VisitExpression(intoExpr.Variable)
            With factory
                Dim _inlineIf_ = .TernaryConditionalExpression(
                                                    .Call(input, DirectCast(.SpecialMember(SpecialMember.System_Nullable_T_get_HasValue), MethodSymbol)),
                                                    .Call(Nothing,
                                                          .WellKnownMember(of MethodSymbol)(WellKnownMember.Microsoft_VisualBasic_CompilerServices_Conversions__ToGenericParameter_T_Object),
                                                          .Call(input, DirectCast(.SpecialMember(SpecialMember.System_Nullable_T_get_Value), MethodSymbol))),
                                                     .Null).MakeRValue
                Dim _passback_ = .ReferenceAssignment(DirectCast(output.ExpressionSymbol,LocalSymbol), _inlineIf_)
                Dim _hasValue_ = .Call( _passback_.MakeRValue, DirectCast(.SpecialMember(SpecialMember.System_Nullable_T_get_HasValue), MethodSymbol))
                Return _hasvalue_.MakeCompilerGenerated
            End With
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
            '    Dim output = TryCast(expression, targetType)
            '    Return output IsNot nothing
            '  End Function
            '

            Dim resultOfTryCast = f.TryCast(Make_DirectCastToObject(f, node.Syntax, typeofexpr.Operand), targettype)

            Dim target = node.Variable
            ' Construct the assignment that passes back the result of the try cast.
            Dim passback = f.AssignmentExpression(target, resultOfTryCast)

            ' Construct the result of passback isnot nothing.
            Dim result = f.ReferenceIsNotNothing(resultOfTryCast)

            ' Finally construct the sequence of code.
            Dim code_sequence = f.Sequence(resultOfTryCast, passback, result)
            Return code_sequence.MakeCompilerGenerated 
        End Function

        Private Function Make_DirectCastToObject(f As SyntheticBoundNodeFactory, node As SyntaxNode, <out> expression As BoundExpression) As BoundExpression
            ' Construct: DirectCast( expression, Object)
            Return New BoundDirectCast(node, expression, Conversions.ClassifyDirectCastConversion(expression.Type, f.[Object], Nothing), f.[Object])
        End Function

        Private Function Rewrite_TypeOf_Into_Variable_With_Structure(f As SyntheticBoundNodeFactory,
                                                                     targettype As TypeSymbol,
                                                                     node As BoundExpressionIntoVariable,
                                                                     typeofexpr As BoundTypeOf,
                                                            Optional isNullable As Boolean = False,
                                                            Optional convKind As ConversionKind = Nothing) As BoundNode
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
            Dim expression_as_object = Make_DirectCastToObject(f, node.Syntax, typeofexpr.Operand).MakeCompilerGenerated

            Dim conv =  Conversions.ClassifyDirectCastConversion(expression_as_object.Type, nullableOf_T, Nothing)
            ' Construct the assignment to the temporary variable with DirectCast( expression_as_object, Nullable(Of T))
            Dim assigment = f.AssignmentExpression(f.Local(tmp_variable, True),
                                             New BoundDirectCast(node.Syntax, expression_as_object, conv, nullableOf_T)).MakeCompilerGenerated
            Dim value =  f.Call(f.Local(tmp_variable, False),
                                                       GetNullableMethod(node.Syntax, nullableOf_T, SpecialMember.System_Nullable_T_GetValueOrDefault)).MakeCompilerGenerated
            ' Construct the assignment that passes back the value or the default value from the temporary variable.
            Dim result As BoundExpression = f.Call(assigment, GetNullableMethod(node.Syntax, nullableOf_T, SpecialMember.System_Nullable_T_get_HasValue)).MakeCompilerGenerated
            Dim passback = f.AssignmentExpression(node.Variable, value.MakeRValue()).MakeCompilerGenerated

            ' Construct the result which check that the temporary variable has value.

            ' Finally construct the sequence of code.
            Dim code_sequence = f.Sequence(tmp_variable, assigment, passback, result)
            Return code_sequence.MakeCompilerGenerated
        End Function

#End Region

    End Class

End Namespace

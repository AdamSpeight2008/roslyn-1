' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports Microsoft.CodeAnalysis.VisualBasic.Symbols

Namespace Microsoft.CodeAnalysis.VisualBasic

  Partial Friend NotInheritable Class LocalRewriter

    Public Overrides Function VisitTypeOfMany(node As BoundTypeOfMany) As BoundNode
      Dim result As BoundExpression= Nothing
      For idx = 0 To node.TargetTypes.Count-1
        Dim expr = node.TargetTypes(idx)
        If idx = 0 Then
           result = expr
        Else If node.IsTypeOfIsNotExpression Then
           result = [AndAlso](expr.Syntax, result, expr)
        Else
           result = [OrElse](expr.Syntax, result, expr)
        End If
      Next
      Return result
    End Function

    Private Function [AndAlso](node As SyntaxNode, lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
      Return New BoundBinaryOperator(node,BinaryOperatorKind.AndAlso, lexpr, rexpr, False,  GetSpecialType(SpecialType.System_Boolean))
    End Function

    Private Function [OrElse](node As SyntaxNode, lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
      Return New BoundBinaryOperator(node,BinaryOperatorKind.OrElse, lexpr, rexpr, False,  GetSpecialType(SpecialType.System_Boolean))
    End Function

  End Class

End Namespace

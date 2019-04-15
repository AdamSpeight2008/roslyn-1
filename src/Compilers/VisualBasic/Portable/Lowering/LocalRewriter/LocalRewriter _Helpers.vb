' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.VisualBasic.Emit
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

        Private Function Make_AND(node As SyntaxNode, LExpe As BoundExpression, RExpr As BoundExpression, RType As TypeSymbol) As BoundExpression
          Return MakeBinaryExpression(node, BinaryOperatorKind.And, LExpe.MakeRValue, RExpr, False, RType)
        End Function

        Private ReadOnly _Boolean_ As NamedTypeSymbol = GetSpecialType(SpecialType.System_Boolean)

        Private Function [AndAlso](node As SyntaxNode, lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
            Return New BoundBinaryOperator(node,BinaryOperatorKind.AndAlso, lexpr, rexpr, False, _Boolean_)
        End Function

        Private Function [OrElse](node As SyntaxNode, lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
            Return New BoundBinaryOperator(node,BinaryOperatorKind.OrElse, lexpr, rexpr, False, _Boolean_)
        End Function

    End Class

End Namespace

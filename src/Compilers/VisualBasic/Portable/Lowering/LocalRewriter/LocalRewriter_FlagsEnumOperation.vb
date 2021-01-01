' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

        Public Overrides Function VisitFlagsEnumOperation(node As BoundFlagsEnumOperation) As BoundNode
            If _inExpressionLambda THen  Return MyBase.VisitFlagsEnumOperation(node)'
            Dim source = VisitExpression(node.Source).MakeRValue()
            Dim original = node.Source.Type.OriginalDefinition
            Dim flagPart = Read_Field(node.Syntax, source, node.FlagName, original)
            Return Rewrite_FlagIsExplicitlySet(node, source, flagPart.MakeRValue)
        End Function

        Private Function Read_Field(node As SyntaxNode,
                               reciever As BoundExpression,
                               fieldSymbol As Symbols.FieldSymbol, type As Symbols.TypeSymbol) As BoundFieldAccess
            Return New BoundFieldAccess(node, reciever, fieldSymbol, False, type).MakeCompilerGenerated()
        End Function

        Private Function Make_AND(node As SyntaxNode, le As BoundExpression, re As BoundExpression) As BoundExpression
           Return MakeBinaryExpression(node, BinaryOperatorKind.And, le, re, False, le.Type)
        End Function

        Private Function Make_EQ(node As SyntaxNode, le As BoundExpression, re As BoundExpression) As BoundExpression
            Return MakeBinaryExpression(node, BinaryOperatorKind.Equals, le, re, False, GetSpecialType(SpecialType.System_Boolean))
        End Function

        Private Function Rewrite_FlagIsExplicitlySet(
                                                      node As BoundFlagsEnumOperation,
                                                      source As BoundExpression,
                                                      flagName As BoundExpression
                                                    ) As BoundNode
            ' <= (source AND flagName ) = flagName
            Dim _And_ = Make_AND(node.Syntax, source, flagName).MakeCompilerGenerated
            Dim _EQ__ = Make_EQ(node.Syntax, _And_, flagName).MakeCompilerGenerated
            Return _EQ__.MakeRValue
        End Function

    End Class

End Namespace

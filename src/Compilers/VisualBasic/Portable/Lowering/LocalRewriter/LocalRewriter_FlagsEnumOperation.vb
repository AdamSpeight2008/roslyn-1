' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

        Public Overrides Function VisitFlagsEnumOperation(node As BoundFlagsEnumOperation) As BoundNode
            If _inExpressionLambda THen  Return MyBase.VisitFlagsEnumOperation(node)'
            Dim source = VisitExpression(node.Source).MakeRValue()
            Dim original = node.Source.Type.OriginalDefinition
            Dim flagPart = (New BoundFieldAccess(node.Syntax, source, node.FlagName, False, original)).MakeRValue()
            Return Rewrite_FlagIsExplicitlySet(node, source, flagPart)
        End Function

        Private Function Rewrite_FlagIsExplicitlySet(
                                                      node As BoundFlagsEnumOperation,
                                                      source As BoundExpression,
                                                      flagName As BoundExpression
                                                    ) As BoundNode
            Dim _And_ = MakeBinaryExpression( node.Syntax,
                            BinaryOperatorKind.And, source, flagName,
                            False, node.Type).MakeCompilerGenerated
            Dim _EQ_ = MakeBinaryExpression(node.Syntax,
                           BinaryOperatorKind.Equals, _AND_.MakeRValue, flagName,
                           False, GetSpecialType(SpecialType.System_Boolean)).MakeCompilerGenerated
            Return _EQ_.MakeRValue
        End Function
    End Class

End Namespace

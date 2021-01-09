' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend NotInheritable Class LocalRewriter

        Public Overrides Function VisitFlagsEnumOperation(node As BoundFlagsEnumOperation) As BoundNode
            If _inExpressionLambda THen  Return MyBase.VisitFlagsEnumOperation(node)'
            Dim source = VisitExpression(node.Source).MakeRValue()
            Dim original = node.Source.Type.OriginalDefinition
            
            If TypeOf node.Flags Is BoundLiteral then
                Dim flagname = node.Flags.ConstantValueOpt.StringValue
                Dim thisFlag As Symbols.FieldSymbol = Nothing
                If Not IsMemberOfThisEnum(original,  flagName, thisFlag) Then
                    Return Nothing
                else
                    Dim flagPart = Read_Field(node.Syntax, source, thisFlag, original)
                    Return AllSpecificFlagsAreEnabled(node, source, flagPart.MakeRValue)
                End if
            Else If TypeOf node.Flags Is BoundExpression Then
                Throw ExceptionUtilities.Unreachable
            End If
                Throw ExceptionUtilities.Unreachable
        End Function

        Private Function IsMemberOfThisEnum(
                                             thisEnumSymbol As Symbols.TypeSymbol,
                                             member As String,
                                       ByRef result As Symbols.FieldSymbol
                                           ) As Boolean
            Dim members = thisEnumSymbol.GetMembers(member.Trim("["c, "]"c))
            result = DirectCast(members.FirstOrDefault(), Symbols.FieldSymbol)
            Return result IsNot Nothing
        End Function
        Private Function Read_Field _
            ( node        As SyntaxNode,
              reciever    As BoundExpression,
              fieldSymbol As Symbols.FieldSymbol,
              type        As Symbols.TypeSymbol
            ) As BoundFieldAccess
            Return New BoundFieldAccess(node, reciever, fieldSymbol, False, type).MakeCompilerGenerated()
        End Function

        Private Function [AND] _
           ( node As SyntaxNode,
             lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
           Return MakeBinaryExpression(node, BinaryOperatorKind.And, lexpr, rexpr, False, lexpr.Type)
        End Function

        Private Function [OR] _
           ( node As SyntaxNode,
             lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
           Return MakeBinaryExpression(node, BinaryOperatorKind.Or, lexpr, rexpr, False, lexpr.Type)
        End Function
        
        Private Function CMP_EQ _
            ( node As SyntaxNode,
              lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
            Return MakeBinaryExpression(node, BinaryOperatorKind.Equals, lexpr, rexpr, False, GetSpecialType(SpecialType.System_Boolean))
        End Function

        Private Function CMP_NE _
            ( node As SyntaxNode,
              lexpr As BoundExpression, rexpr As BoundExpression) As BoundExpression
            Return MakeBinaryExpression(node, BinaryOperatorKind.NotEquals, lexpr, rexpr, False, GetSpecialType(SpecialType.System_Boolean))
        End Function

        Private Function [NOT]( expr As BoundExpression) As BoundExpression
            Return New BoundUnaryOperator(expr.Syntax, UnaryOperatorKind.Not, expr, False, expr.Type)
        End Function

        Private Function Make_Zero(node As SyntaxNode) As BoundExpression
            Return New BoundLiteral(node, ConstantValue.Nothing, GetSpecialType(SpecialType.System_Int32))
        End Function

        Private Function DisableSpecificFlags _
            ( node   As BoundFlagsEnumOperation,
              source As BoundExpression,
              flags  As BoundExpression
            ) As BoundNode
            ' <= (source AND (NOT flagName ))
            Return [AND](node.Syntax, source, [NOT](flags)).MakeCompilerGenerated.MakeRValue 
        End Function
        Private Function EnableSpecificFlags _
            ( node   As BoundFlagsEnumOperation,
              source As BoundExpression,
              flags  As BoundExpression
            ) As BoundNode
            ' <= (source AND flagName )
            Return [OR](node.Syntax, source, flags).MakeCompilerGenerated.MakeRValue
        End Function

        Private Function AnySpecificFlagsAreEnabled _
            ( node   As BoundFlagsEnumOperation,
              source As BoundExpression,
              flags  As BoundExpression
            ) As BoundNode
            ' <= (source AND flagName ) <> 0
            Dim _And_ = [AND](node.Syntax, source, flags).MakeCompilerGenerated
            Dim _NE__ = CMP_NE(node.Syntax, _And_, Make_Zero(node.Syntax)).MakeCompilerGenerated
            Return _NE__.MakeRValue
        End Function

        Private Function AllSpecificFlagsAreEnabled _
            ( node   As BoundFlagsEnumOperation,
              source As BoundExpression,
              flags  As BoundExpression
            ) As BoundNode
            ' <= (source AND flagName ) = flagName
            Dim _And_ = [AND](node.Syntax, source, flags).MakeCompilerGenerated
            Dim _EQ__ = CMP_EQ(node.Syntax, _And_, flags).MakeCompilerGenerated
            Return _EQ__.MakeRValue
        End Function

    End Class

End Namespace

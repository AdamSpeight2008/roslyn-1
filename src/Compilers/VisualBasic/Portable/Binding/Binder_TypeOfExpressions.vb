' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Microsoft.CodeAnalysis.VisualBasic
      
    Partial Friend Class Binder

        Private Function BindTypeOfExpression _ 
            (
              node          As TypeOfExpressionSyntax,
              diagnostics   As DiagnosticBag
            ) As BoundExpression
            Debug.Assert(node.Kind.IsEither(SyntaxKind.TypeOfIsExpression, SyntaxKind.TypeOfIsNotExpression))
            Dim operand         = BindRValue(node.Expression, diagnostics, isOperandOfConditionalBranch:=False)
            Dim operandType     = operand.Type
            Dim operatorIsIsNot = (node.Kind = SyntaxKind.TypeOfIsNotExpression)
            Return BindTypeOf_Inner(node, diagnostics, operand, operandType, operatorIsIsNot, node.Type, node.Expression, True)
        End Function

        Private Function BindTypeOf_Inner _ 
            (
              node              As SyntaxNode,
        ByRef diagnostics       As DiagnosticBag,
        ByRef operand           As BoundExpression,
              operandType       As TypeSymbol,
              operatorIsIsNot   As Boolean,
              targetTypeSyntax  As TypeSyntax,
              expr              As ExpressionSyntax,
     Optional suppressAdditionalDiagostics As Boolean = False
            ) As BoundTypeOf

          Dim resultType   As TypeSymbol  = GetSpecialType(SpecialType.System_Boolean, node, diagnostics)
          Dim targetSymbol As Symbol      = BindTypeOrAliasSyntax(targetTypeSyntax, diagnostics)
          Dim targetType = DirectCast(If(TryCast(targetSymbol, TypeSymbol), DirectCast(targetSymbol, AliasSymbol).Target), TypeSymbol)
          Dim btExpr As New BoundTypeExpression(targetTypeSyntax, targetType)

          If operand.HasErrors OrElse operandType.IsErrorType() OrElse targetType.IsErrorType() Then
             ' If operand is bad or either the source or target types have errors, bail out preventing more cascading errors.
             Return New BoundTypeOf(node, btExpr.Type, operand, operatorIsIsNot, resultType)
          End If

          If Not operandType.IsReferenceType AndAlso Not operandType.IsTypeParameter() Then
             ReportDiagnostic(diagnostics, expr, ERRID.ERR_TypeOfRequiresReferenceType1, operandType)
          Else
             Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
             Dim convKind           As ConversionKind = Conversions.ClassifyTryCastConversion(operandType, targetType, useSiteDiagnostics)
             If diagnostics.Add(node, useSiteDiagnostics) Then
                If suppressAdditionalDiagostics Then diagnostics = New DiagnosticBag() ' Suppress any additional diagnostics
             ElseIf Not Conversions.ConversionExists(convKind) Then
                diagnostics.Add(ERRID.ERR_TypeOfExprAlwaysFalse2, node.Location, operandType, targetType)
             End If
          End If

          If operandType.IsTypeParameter() Then
             Dim obj = GetSpecialType(SpecialType.System_Object, node, diagnostics)
             operand = ApplyImplicitConversion(node, obj, operand, diagnostics)
          End If
          Return New BoundTypeOf(node, targetType, operand, operatorIsIsNot, resultType)
        End Function

        Private Function BindTypeOfManyExpression _ 
            (
              node          As TypeOfManyExpressionSyntax,
              diagnostics   As DiagnosticBag
            ) As BoundTypeOfBase

            Debug.Assert(node.Kind.IsEither(SyntaxKind.TypeOfManyIsExpression,SyntaxKind.TypeOfManyIsNotExpression))

            Dim operand         = BindRValue(node.Expression, diagnostics, isOperandOfConditionalBranch:=False)
            Dim operandType     = operand.Type
            Dim operatorIsIsNot = (node.Kind = SyntaxKind.TypeOfManyIsNotExpression)
            Dim targertTypes    = ImmutableArray(Of BoundTypeOf).Empty
            Dim resultType As TypeSymbol = GetSpecialType(SpecialType.System_Boolean, node, diagnostics)

            For Each _type_ In node.Types

              Dim _operand_         = BindRValue(node.Expression, diagnostics, isOperandOfConditionalBranch:=False).MakeCompilerGenerated
              Dim innerTypeOfExpr   = BindTypeOf_Inner(_type_, diagnostics, _operand_, operandType, operatorIsIsNot, _type_, node.Expression)

              If innerTypeOfExpr.HasErrors Then ReportDiagnostic(diagnostics, _type_.GetDiagnostics.FirstOrDefault)
              targertTypes = targertTypes.Add(innerTypeOfExpr.MakeCompilerGenerated)

            Next
            Return New BoundTypeOfMany(node, targertTypes, operand, operatorIsIsNot, resultType)
        End Function

    End Class

End Namespace

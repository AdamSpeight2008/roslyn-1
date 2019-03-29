' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic.LanguageFeatures.CheckFeatureAvailability

Namespace Microsoft.CodeAnalysis.VisualBasic
      
    Partial Friend Class Binder
        
        Private Shared Sub ReportNoDefaultProperty(expr As BoundExpression, diagnostics As DiagnosticBag)
            Dim type = expr.Type
            Dim syntax = expr.Syntax
            Select Case type.TypeKind
                Case TypeKind.Class     : ReportDiagnostic(diagnostics, syntax, ERRID.ERR_NoDefaultNotExtend1, type) ' "Class '{0}' cannot be indexed because it has no default property."
                Case TypeKind.Structure : ReportDiagnostic(diagnostics, syntax, ERRID.ERR_StructureNoDefault1, type) ' "Structure '{0}' cannot be indexed because it has no default property."
                Case TypeKind.Error     : ' We should have reported an error elsewhere.
                Case Else               : ReportDiagnostic(diagnostics, syntax, ERRID.ERR_InterfaceNoDefault1, type) ' "'{0}' cannot be indexed because it has no default property."
            End Select                   
        End Sub

        Private Shared Sub ReportQualNotObjectRecord(expr As BoundExpression, diagnostics As DiagnosticBag)
            ' "'!' requires its left operand to have a type parameter, class or interface type, but this operand has the type '{0}'."
            ReportDiagnostic(diagnostics, expr.Syntax, ERRID.ERR_QualNotObjectRecord1, expr.Type)
        End Sub

        Private Shared Sub ReportDefaultMemberNotProperty(expr As BoundExpression, diagnostics As DiagnosticBag)
            ' "Default member '{0}' is not a property."
            ' Note: The error argument is the expression type
            ' rather than the expression text used in Dev10.
            ReportDiagnostic(diagnostics, expr.Syntax, ERRID.ERR_DefaultMemberNotProperty1, expr.Type)
        End Sub

        Private Function BindDictionaryAccess(node As MemberAccessExpressionSyntax, diagnostics As DiagnosticBag) As BoundExpression
            Debug.Assert(node.OperatorToken.Kind = SyntaxKind.ExclamationToken)
            Dim leftOpt = node.Expression
            Dim left As BoundExpression
            If leftOpt Is Nothing Then
                ' Spec 11.7: "If an exclamation point is specified with no expression, the
                ' expression from the immediately containing With statement is assumed.
                ' If there is no containing With statement, a compile-time error occurs."

                Dim conditionalAccess As ConditionalAccessExpressionSyntax = node.GetCorrespondingConditionalAccessExpression()

                If conditionalAccess IsNot Nothing Then
                    left = GetConditionalAccessReceiver(conditionalAccess)
                Else
                    left = TryBindOmittedLeftForDictionaryAccess(node, Me, diagnostics)
                End If

                If left Is Nothing Then
                    ' Didn't find binder that can handle member access with omitted left part

                    Return BadExpression(node, ImmutableArray.Create(
                                          ReportDiagnosticAndProduceBadExpression(diagnostics, node, ERRID.ERR_BadWithRef),
                                          New BoundLiteral(
                                                node.Name,
                                                ConstantValue.Create(node.Name.ToString),
                                                GetSpecialType(SpecialType.System_String, node.Name, diagnostics))),
                                          ErrorTypeSymbol.UnknownResultType)
                End If
            Else
                left = BindExpression(leftOpt, diagnostics)
            End If

            If Not left.IsLValue AndAlso left.Kind <> BoundKind.LateMemberAccess Then
                left = MakeRValue(left, diagnostics)
                Debug.Assert(left IsNot Nothing)
            End If

            Dim type = left.Type
            Debug.Assert(type IsNot Nothing)

            If Not type.IsErrorType() Then
                Dim nameSyntax = TryCast(node.Name, SimpleNameSyntax)
                If nameSyntax IsNot Nothing Then
                    ' (left : Objext)!( ___ )
                    If type.SpecialType = SpecialType.System_Object OrElse type.IsExtensibleInterfaceNoUseSiteDiagnostics() Then
                        Return BindLeftSideAsObject(node, diagnostics, left, nameSyntax)
                    End If

                    ' (left : Interface)!( ___ )
                    If type.IsInterfaceType Then
                        BindLeftSideAsInterface(node, diagnostics, type)
                    End If

                    Dim defaultPropertyGroup As BoundExpression = BindDefaultPropertyGroup(node, left, diagnostics)
                    Debug.Assert(defaultPropertyGroup Is Nothing OrElse defaultPropertyGroup.Kind = BoundKind.PropertyGroup OrElse
                             defaultPropertyGroup.Kind = BoundKind.MethodGroup OrElse defaultPropertyGroup.HasErrors)

                    ' Dev10 limits Dictionary access to properties.
                    If defaultPropertyGroup IsNot Nothing AndAlso defaultPropertyGroup.Kind = BoundKind.PropertyGroup Then
                        Dim name = node.Name
                        Dim arg = New BoundLiteral(name, ConstantValue.Create(nameSyntax.Identifier.ValueText), GetSpecialType(SpecialType.System_String, name, diagnostics))
                        Return BindInvocationExpression(node,
                                                     left.Syntax,
                                                     TypeCharacter.None,
                                                     DirectCast(defaultPropertyGroup, BoundPropertyGroup),
                                                     boundArguments:=ImmutableArray.Create(Of BoundExpression)(arg),
                                                      argumentNames:=Nothing,
                                                        diagnostics:=diagnostics,
                                              isDefaultMemberAccess:=True,
                                                      callerInfoOpt:=node)

                    ElseIf defaultPropertyGroup Is Nothing OrElse Not defaultPropertyGroup.HasErrors Then

                        Select Case type.TypeKind
                            Case TypeKind.Enum
                                If InternalSyntax.Feature.EnumFlagOperators.CheckFeatureAvailability(
                                    DirectCast(left.Syntax.SyntaxTree.Options,VisualBasicParseOptions),
                                    node.OperatorToken.GetLocation(),diagnostics) Then

                                    Dim original = type.OriginalDefinition
                                    ' Make sure the enum has the <Flags> attribute.
                                    If IsFlagsEnum(DirectCast(original, INamedTypeSymbol)) Then
                                        Return BindEnumFlagExpression(original, node, left, node, diagnostics)
                                    Else
                                        Return ReportDiagnosticAndProduceBadExpression(diagnostics, node, ERRID.ERR_MissingFlagsAttributeOnEnum, original.Name)
                                    End If
                                Else
                                    ReportQualNotObjectRecord(left, diagnostics)
                                End If

                            Case TypeKind.Array
                                ReportQualNotObjectRecord(left, diagnostics)
                            Case TypeKind.Class
                                If type.SpecialType = SpecialType.System_Array Then
                                    ReportDefaultMemberNotProperty(left, diagnostics)
                                Else
                                    ReportNoDefaultProperty(left, diagnostics)
                                End If
                            Case TypeKind.TypeParameter,
                                 TypeKind.Interface         : ReportNoDefaultProperty(left, diagnostics)
                            Case TypeKind.Structure
                                If type.IsIntrinsicValueType() Then
                                    ReportQualNotObjectRecord(left, diagnostics)
                                Else
                                    ReportNoDefaultProperty(left, diagnostics)
                                End If
                            Case Else
                                ReportQualNotObjectRecord(left, diagnostics)
                                ReportDefaultMemberNotProperty(left, diagnostics)
                        End Select
                    End If
                End If
            End If

            Return BadExpression(node,
                                 ImmutableArray.Create(left,
                                                       New BoundLiteral(node.Name,
                                                                        ConstantValue.Create(node.Name.ToString),
                                                      GetSpecialType(SpecialType.System_String, node.Name, diagnostics))),
                                 ErrorTypeSymbol.UnknownResultType)
        End Function

        Private Shared Sub BindLeftSideAsInterface(node As MemberAccessExpressionSyntax, diagnostics As DiagnosticBag, type As TypeSymbol)
            Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
            ' In case IsExtensibleInterfaceNoUseSiteDiagnostics above failed because there were bad inherited interfaces.
            type.AllInterfacesWithDefinitionUseSiteDiagnostics(useSiteDiagnostics)
            diagnostics.Add(node, useSiteDiagnostics)
        End Sub

        Private Function BindLeftSideAsObject(node As MemberAccessExpressionSyntax, diagnostics As DiagnosticBag, left As BoundExpression, nameSyntax As SimpleNameSyntax) As BoundExpression
            Dim name = nameSyntax.Identifier
            Dim arg = New BoundLiteral(nameSyntax, ConstantValue.Create(name.ValueText), GetSpecialType(SpecialType.System_String, name, diagnostics))
            Dim boundArguments = ImmutableArray.Create(Of BoundExpression)(arg)
            Return BindLateBoundInvocation(node, Nothing, left, boundArguments, Nothing, diagnostics)
        End Function

        Private Const _FlagsAttribute_ = "FlagsAttribute"
        Private Const _System_ = "System"

        Private Shared ReadOnly IsFlagsAttribute As Func(Of AttributeData, Boolean) =
            Function(attribute)
                Dim ctor = attribute.AttributeConstructor
                If (ctor Is Nothing) Then Return False
                Dim [Type] = ctor.ContainingType
                If (ctor.Parameters.Any() OrElse [Type].Name <> _FlagsAttribute_) Then Return False
                Dim containingSymbol = Type.ContainingSymbol
                Return (containingSymbol.Kind = SymbolKind.Namespace) AndAlso
                       (containingSymbol.Name = _System_) AndAlso DirectCast(containingSymbol.ContainingSymbol, INamespaceSymbol).IsGlobalNamespace
            End Function

        Private Function IsFlagsEnum(typeSymbol As INamedTypeSymbol) As Boolean
            Return (typeSymbol.TypeKind = TypeKind.Enum) AndAlso typeSymbol.GetAttributes().Any(IsFlagsAttribute)
        End Function

        Private Function IsMemberOfThisEnum(
                                             thisEnumSymbol As TypeSymbol,
                                             member As String,
                                       ByRef result As FieldSymbol
                                           ) As Boolean
            Dim members = thisEnumSymbol.GetMembers(member)
            result = DirectCast(members.FirstOrDefault(), FieldSymbol)
            Return result IsNot Nothing
        End Function

        Friend Function BindEnumFlagExpression(
                                                original As TypeSymbol,
                                                node As SyntaxNode,
                                                expr As BoundExpression,
                                                member As MemberAccessExpressionSyntax,
                                                diagBag As DiagnosticBag
                                              ) As BoundExpression
            Dim EnumMember = TryCast(member.Expression, SimpleNameSyntax)
            If EnumMember IsNot Nothing Then
                Dim FlagName = EnumMember.Identifier.ValueText
                Return Bind_FlagsEnumOperation_WithEnumMember(FlagName, member, EnumMember, member.OperatorToken, EnumMember, diagBag)
            ElseIf member.Name IsNot Nothing Then
                Return Bind_FlagsEnumOperation_WithExpression(member, member.Expression, member.OperatorToken, member.Name, diagBag)
            Else
                Throw ExceptionUtilities.Unreachable()
                Return Nothing
            End If
        End Function

        Private Function GetFlagsEnumOperationKind(
                                                    node As SyntaxToken
                                                  ) As FlagsEnumOperatorKind
            Select Case node.Kind
                Case SyntaxKind.FlagsEnumClearToken     : Return FlagsEnumOperatorKind.Clear
                Case SyntaxKind.FlagsEnumIsSetToken,    
                     SyntaxKind.ExclamationToken        : Return FlagsEnumOperatorKind.IsSet
                Case SyntaxKind.FlagsEnumSetToken       : Return FlagsEnumOperatorKind.Set
                Case SyntaxKind.FlagsEnumIsAnyToken     : Return FlagsEnumOperatorKind.IsAny
                Case Else                               : Return FlagsEnumOperatorKind.None
            End Select
        End Function

        Private Function Bind_FlagsEnumOperation_WithEnumMember(
                                                                 FlagName As String,
                                                                 node As ExpressionSyntax,
                                                                 EnumFlags As ExpressionSyntax,
                                                                 opToken As SyntaxToken,
                                                                 EnumFlag As SimpleNameSyntax,
                                                                 diagBag As DiagnosticBag
                                                               ) As BoundExpression
            Debug.Assert(EnumFlags isnot nothing)
            Dim eFlag As FieldSymbol = Nothing
            Dim bFlags = BindExpression(EnumFlags, diagBag)
            Dim original = bFlags.Type.OriginalDefinition

            If Not IsFlagsEnum(DirectCast(original, INamedTypeSymbol)) Then Return ReportDiagnosticAndProduceBadExpression(diagBag, EnumFlags, ERRID.ERR_EnumNotExpression1, original.Name)
            If Not IsMemberOfThisEnum(original, FlagName, eFlag)       Then Return ReportDiagnosticAndProduceBadExpression(diagBag, EnumFlag, ERRID.ERR_NameNotMember2, FlagName, original.Name)

            Dim opKind = GetFlagsEnumOperationKind(opToken)
            If opKind = FlagsEnumOperatorKind.None                     Then Return ReportDiagnosticAndProduceBadExpression(diagBag, 
                                                                                                                           DirectCast(node.GetNodeSlot(1), VisualBasicSyntaxNode), 
                                                                                                                           ERRID.ERR_UnknownOperator, opKind.ToString, original.Name)

            Return New BoundFlagsEnumOperationExpressionSyntax(
                  syntax:= node,
               enumFlags:= bFlags, op:=opKind,
                enumFlag:= New BoundFieldAccess(EnumFlag, bFlags, eFlag, False, original),
                    type:= If(opKind = FlagsEnumOperatorKind.IsSet Or opKind = FlagsEnumOperatorKind.IsAny,
                                  GetSpecialType(SpecialType.System_Boolean, EnumFlag, diagBag), original)
                    )
        End Function

        Private Function Bind_FlagsEnumOperation(
                                                  node As FlagsEnumOperationExpressionSyntax,
                                                  diagBag As DiagnosticBag
                                                ) As BoundExpression
            If node.EnumFlags Is nothing Then Return ReportDiagnosticAndProduceBadExpression(diagBag, node, ERRID.ERR_ExpectedExpression)
            Dim identifer = TryCast(node.EnumFlag, SimpleNameSyntax)
            If identifer Is Nothing Then Return Bind_FlagsEnumOperation_WithExpression(node, node.EnumFlags, node.OperatorToken, node.EnumFlag, diagBag)
            Dim FlagName = If(identifer.Identifier.ValueText, String.Empty)
            Return Bind_FlagsEnumOperation_WithEnumMember(FlagName, node, node.EnumFlags, node.OperatorToken, identifer, diagBag)
        End Function

        Private Function Bind_FlagsEnumOperation_WithExpression(
                                                                 node As ExpressionSyntax,
                                                                 enumFlags As ExpressionSyntax,
                                                                 opToken As SyntaxToken,
                                                                 enumFlag As ExpressionSyntax,
                                                                 diagBag As DiagnosticBag
                                                               ) As BoundExpression
            Dim eFlag As FieldSymbol = Nothing
            Dim bFlags = BindExpression(enumFlags, diagBag)
            Dim original = bFlags.Type.OriginalDefinition

            If Not IsFlagsEnum(DirectCast(original, INamedTypeSymbol)) Then
                Return ReportDiagnosticAndProduceBadExpression(diagBag, enumFlags, ERRID.ERR_EnumNotExpression1)
            End If

            Dim opKind = GetFlagsEnumOperationKind(opToken)
            If opKind = FlagsEnumOperatorKind.None Then
                Return ReportDiagnosticAndProduceBadExpression(diagBag, DirectCast(node.GetNodeSlot(1), VisualBasicSyntaxNode), ERRID.ERR_UnknownOperator, opKind.ToString, original.Name)
            End If
            If TypeOf enumFlag IsNot ParenthesizedExpressionSyntax Then
                Return ReportDiagnosticAndProduceBadExpression(diagBag, enumFlag, ERRID.ERR_ExpectedParenthesizedExpression, opKind.ToString, original.Name)
            End If
            Return New BoundFlagsEnumOperationExpressionSyntax(
                  syntax:= node,
               enumFlags:= bFlags, opKind,
                enumFlag:= BindExpression(enumFlag, diagBag),
                    type:= If(opKind = FlagsEnumOperatorKind.IsSet Or opKind = FlagsEnumOperatorKind.IsAny,
                            GetSpecialType(SpecialType.System_Boolean, enumFlag, diagBag), original)
                    )
        End Function

      End Class
End Namespace

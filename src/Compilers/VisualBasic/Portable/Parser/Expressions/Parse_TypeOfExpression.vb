' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

'
'============ Methods for parsing portions of executable statements ==
'
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.Syntax.InternalSyntax
Imports InternalSyntaxFactory = Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax.SyntaxFactory

Namespace Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax
    
    Partial Friend Class Parser

         ''' <summary>
        ''' Parse TypeOf ... Is ... or TypeOf ... IsNot ...
        ''' TypeOfExpression -> "TypeOf" Expression "Is|IsNot" LineTerminator? TypeName
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ParseTypeOf() As TypeOfExpressionSyntax
            Debug.Assert(CurrentToken.Kind = SyntaxKind.TypeOfKeyword, "must be at TypeOf.")
            ' TypeOfExpression <: Expression ::= TypeOfKeyword Expression IsTypeClauseSyntax

            Dim [typeOf] = Parse_Keyword(SyntaxKind.TypeOfKeyword) ' Consume 'TypeOf'.
            Dim operand  = ParseExpressionCore(OperatorPrecedence.PrecedenceRelational) 'Dev10 uses ParseVariable

            If operand.ContainsDiagnostics Then operand = ResyncAt(operand, SyntaxKind.IsKeyword, SyntaxKind.IsNotKeyword)
            Dim kind = SyntaxKind.TypeOfIsExpression
            Dim isTypeClause = Parse_IsTypeClause()
            Select Case isTypeClause.Kind
                   Case SyntaxKind.IsTypeClause    : kind = SyntaxKind.TypeOfIsExpression
                   Case SyntaxKind.IsNotTypeClause : kind = SyntaxKind.TypeOfIsNotExpression
                   Case Else : Throw ExceptionUtilities.UnexpectedValue(kind)
            End Select

            Return SyntaxFactory.TypeOfExpression(kind, [typeOf], operand, isTypeClause)
        End Function

        Private Function Parse_KindOfIsKeyword() As KeywordSyntax
            ' KindOfIsKeyword ::= IsKeyword | IsNotKeyword ;;
            Dim isKeyword As KeywordSyntax = Nothing
            Select Case True
                   Case TryParse_Keyword(SyntaxKind.IsKeyword, isKeyword)     ' Is is supported
                   Case TryParse_Keyword(SyntaxKind.IsNotKeyword, isKeyword)  ' ISNOT  is conditionally supported 
                        isKeyword = CheckFeatureAvailability(Feature.TypeOfIsNot, isKeyword)
                   Case Else
                        isKeyword = DirectCast(HandleUnexpectedToken(SyntaxKind.IsKeyword), KeywordSyntax)
            End Select
            Return isKeyword
        End Function

        Private Function Parse_IsTypeClause() As IsTypeClauseSyntax
            ' IsTypeClause ::= KindOfIsKeyword DeclarationAsClause? KindOfTypeTarget
            Dim operatorToken = Parse_KindOfIsKeyword()
            Dim optionalDeclarationAsClause = Parse_DeclarationAsClause

            Dim targetType = ParseKindOfTypeTarget()

            Validate_IsTypeClause(operatorToken, optionalDeclarationAsClause, targetType)
            Select Case operatorToken.Kind
                Case SyntaxKind.IsKeyword    : Return SyntaxFactory.IsTypeClause(operatorToken, optionalDeclarationAsClause, targetType)
                Case SyntaxKind.IsNotKeyword : Return SyntaxFactory.IsNotTypeClause(operatorToken, optionalDeclarationAsClause, targetType)
            End Select
            Throw ExceptionUtilities.Unreachable()
        End Function

        Private Sub Validate_IsTypeClause(
                                     ByRef operatorToken As KeywordSyntax,
                                     BYRef optionalDeclarationAsClause As DeclarationAsClauseSyntax,
                                     ByRef targetType As VisualBasicSyntaxNode
                                         )
            If optionalDeclarationAsClause Is Nothing Then Exit Sub
            If targetType.Kind = SyntaxKind.TypeArgumentList Then _
                optionalDeclarationAsClause = optionalDeclarationAsClause.AddError(ERRID.ERR_DeclarationAsClauseDoesNotSupportTypeArgumentList)

            If operatorToken.Kind = SyntaxKind.IsNotKeyword Then _
                optionalDeclarationAsClause = optionalDeclarationAsClause.AddError(ERRID.ERR_DeclarationAsClauseDoesNotSupportIsNotOperator)
        End Sub

        Private Function ParseKindOfTypeTarget() As VisualBasicSyntaxNode
            ' KindOfType ::= TypeArgumentListSyntax | TypeSyntax ;;
            Dim targetType As VisualBasicSyntaxNode
            If IsAtStartOfArgumentTypeList THen
                targetType = ParseGenericArguments(True, true)
                targetType = CheckFeatureAvailability( Feature.TypeOfMany, targetType)
            Else
                targetType = ParseGeneralType()
            End If
            Return targetType
        End Function

        Private Function IsPossible_NameAs() As Boolean
            Return PeekNextToken.Kind = SyntaxKind.AsKeyword
        End Function

        Private Function Parse_DeclarationAsClause( ) As DeclarationAsClauseSyntax
            If IsPossible_NameAs = False Then Return Nothing
            Dim declaration = ParseIdentifierAllowingKeyword()
            Dim as_keyword = Parse_Keyword(SyntaxKind.AsKeyword, eatLine:= True)
            Return  CheckFeatureAvailability(Feature.TypeOfAs, SyntaxFactory.DeclarationAsClause(declaration, as_keyword))
        End Function
    
    End Class

End Namespace

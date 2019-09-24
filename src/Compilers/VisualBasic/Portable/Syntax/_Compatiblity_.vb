Namespace Microsoft.CodeAnalysis.VisualBasic
    Namespace Syntax
        Partial Public Class TypeOfExpressionSyntax

            Public Function Update(kind As SyntaxKind, typeOfKeyword As SyntaxToken, expression As ExpressionSyntax, operatorToken As SyntaxToken, type As TypeSyntax) As TypeOfExpressionSyntax
                Return Me.Update(kind, typeOfKeyword, expression, operatorToken, type, nothing)
            End Function

        End Class

    End Namespace

    Partial Public Class SyntaxFactory
        Public Shared Function TypeOfExpression(kind As SyntaxKind, typeOfKeyword As SyntaxToken, expression As Syntax.ExpressionSyntax, operatorToken As SyntaxToken, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
            Return SyntaxFactory.TypeOfExpression(kind, typeOfKeyword, expression, operatorToken, type, nothing)
        End Function

        'Public Shared Function TypeOfExpression(kind As SyntaxKind, expression As Syntax.ExpressionSyntax, operatorToken As SyntaxToken, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
        '    Return SyntaxFactory.TypeOfExpression(kind, expression, operatorToken, type, Nothing)
        'End Function
        'Public Shared Function TypeOfIsExpression(expression As Syntax.ExpressionSyntax, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
        '    Return SyntaxFactory.TypeOfIsExpression(expression, type, nothing)
        'End Function
        Public Shared Function TypeOfIsExpression(typeOfKeyword As SyntaxToken, expression As Syntax.ExpressionSyntax, operatorToken As SyntaxToken, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
            Return TypeOfIsExpression(typeOfKeyword, expression, operatorToken, type, nothing)
        End Function

        'Public Shared Function TypeOfIsNotExpression(expression As Syntax.ExpressionSyntax, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
        '    Return SyntaxFactory.TypeOfIsNotExpression(expression, type, nothing)
        'End Function

        Public Shared Function TypeOfIsNotExpression(typeOfKeyword As SyntaxToken, expression As Syntax.ExpressionSyntax, operatorToken As SyntaxToken, type As Syntax.TypeSyntax) As Syntax.TypeOfExpressionSyntax
            Return SyntaxFactory.TypeOfIsNotExpression(typeOfKeyword, expression, operatorToken, type, nothing)
        End Function

    End Class

End Namespace

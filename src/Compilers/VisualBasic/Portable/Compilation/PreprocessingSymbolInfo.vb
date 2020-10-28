' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Roslyn.Utilities

Namespace Microsoft.CodeAnalysis.VisualBasic
    Friend Structure VisualBasicPreprocessingSymbolInfo
        Implements IEquatable(Of VisualBasicPreprocessingSymbolInfo)

        Friend Shared None As New VisualBasicPreprocessingSymbolInfo(Nothing, Nothing, False)

        ''' <summary> The symbol that was referred to by the identifier, if any. </summary>
        Public ReadOnly Property Symbol As PreprocessingSymbol

        Public Property IsDefined As Boolean

        ''' <summary> Returns the constant value associated with the symbol, if any. </summary>
        Public ReadOnly Property ConstantValue As Object

        Public Shared Widening Operator CType(info As VisualBasicPreprocessingSymbolInfo) As PreprocessingSymbolInfo
            Return New PreprocessingSymbolInfo(info.Symbol, info.IsDefined)
        End Operator

        Friend Sub New( symbol           As PreprocessingSymbol,
                        constantValueOpt As Object,
                        isDefined        As Boolean
                      )
            Me.Symbol = symbol
            Me.ConstantValue = constantValueOpt
            Me.IsDefined = isDefined
        End Sub

        Public Overloads Function Equals(other As VisualBasicPreprocessingSymbolInfo) As Boolean Implements IEquatable(Of VisualBasicPreprocessingSymbolInfo).Equals
            Return IsDefined = other.IsDefined AndAlso _symbol = other._symbol AndAlso
                   Equals(Me._constantValue, other._constantValue)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return TypeOf obj Is VisualBasicTypeInfo AndAlso
                   Equals(DirectCast(obj, VisualBasicTypeInfo))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Hash.Combine(_symbol, Hash.Combine(_constantValue, CInt(IsDefined)))
        End Function

    End Structure

End Namespace

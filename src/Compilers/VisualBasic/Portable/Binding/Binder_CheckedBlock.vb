' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports TypeKind = Microsoft.CodeAnalysis.TypeKind

Namespace Microsoft.CodeAnalysis.VisualBasic

    Friend NotInheritable Class CheckedBlockBinder
        Inherits Binder

        Private ReadOnly _checkedBlockSyntax As CheckedBlockSyntax
        ''' <summary> Create a new instance of Checked Block statement binder for a statement syntax provided </summary>
        Public Sub New(enclosing As Binder, syntax As CheckedBlockSyntax)
            MyBase.New(enclosing)

            Debug.Assert(syntax IsNot Nothing)
            Debug.Assert(syntax.BeginCheckedBlockStatement IsNot Nothing)
            Me._checkedBlockSyntax = syntax
        End Sub

        'Friend Overrides ReadOnly Property Locals As ImmutableArray(Of LocalSymbol)
        '    Get
        '        Return ImmutableArray(Of LocalSymbol).Empty
        '    End Get
        'End Property

        Friend Overrides Function BindVariableDeclaration(tree As VisualBasicSyntaxNode,
                                                          name As ModifiedIdentifierSyntax,
                                                          asClauseOpt As AsClauseSyntax,
                                                          equalsValueOpt As EqualsValueSyntax,
                                                          diagnostics As DiagnosticBag,
                                                 Optional skipAsNewInitializer As Boolean = False) As BoundLocalDeclaration
            Return Me.m_containingBinder.BindVariableDeclaration(tree, name, asClauseOpt, equalsValueOpt, diagnostics, skipAsNewInitializer)
        End Function

    End Class

End Namespace


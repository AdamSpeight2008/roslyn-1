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
Imports Microsoft.CodeAnalysis.VisualBasic.BinderFlagsExtensions

Namespace Microsoft.CodeAnalysis.VisualBasic

    Friend NotInheritable Class CheckedBlockBinder
        Inherits Binder

        Private ReadOnly _checkedBlockSyntax As CheckedBlockSyntax
        ''' <summary> Create a new instance of Checked Block statement binder for a statement syntax provided </summary>
        Public Sub New(enclosing As Binder, syntax As CheckedBlockSyntax)
            MyBase.New(enclosing, UpdateCheckRegionFlags(enclosing.Flags, DecodeOnOff(syntax.BeginCheckedBlockStatement.OnOrOffKeyword)))

            Debug.Assert(syntax IsNot Nothing)
            Debug.Assert(syntax.BeginCheckedBlockStatement IsNot Nothing)
            Me._checkedBlockSyntax = syntax
        End Sub

        Private Shared Function UpdateCheckRegionFlags(f As BinderFlags, checkIntegerOverflow As Boolean) As BinderFlags
            Dim added As BinderFlags = If(checkIntegerOverflow, BinderFlags.CheckedRegion, BinderFlags.UncheckedRegion)
            Dim removed As BinderFlags = If(checkIntegerOverflow, BinderFlags.UncheckedRegion, BinderFlags.CheckedRegion)

            If f.Includes(added) Then
                Return f
            End If
            Return ((f And Not removed) Or added)
        End Function

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


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

    Friend NotInheritable Class OverflowHandlingBinder
        Inherits Binder

        ''' <summary> Create a new instance of Checked Block statement binder for a statement syntax provided </summary>
        Public Sub New(enclosing As Binder, checkOverflow As Boolean)
            MyBase.New(enclosing, UpdateCheckRegionFlags(enclosing.Flags, checkOverflow))
        End Sub

        Private Shared Function UpdateCheckRegionFlags(f As BinderFlags, checkIntegerOverflow As Boolean) As BinderFlags
            Dim added As BinderFlags = If(checkIntegerOverflow, BinderFlags.CheckedRegion, BinderFlags.UncheckedRegion)
            Dim removed As BinderFlags = If(checkIntegerOverflow, BinderFlags.UncheckedRegion, BinderFlags.CheckedRegion)
            Return f.AddAndRemove(added, removed)
        End Function

        'Friend Overrides Function BindVariableDeclaration(tree As VisualBasicSyntaxNode,
        '                                                  name As ModifiedIdentifierSyntax,
        '                                                  asClauseOpt As AsClauseSyntax,
        '                                                  equalsValueOpt As EqualsValueSyntax,
        '                                                  diagnostics As DiagnosticBag,
        '                                         Optional skipAsNewInitializer As Boolean = False) As BoundLocalDeclaration
        '    Return Me.m_containingBinder.BindVariableDeclaration(tree, name, asClauseOpt, equalsValueOpt, diagnostics, skipAsNewInitializer)
        'End Function

        'Public Overrides Function DeclareImplicitLocalVariable(nameSyntax As IdentifierNameSyntax, diagnostics As DiagnosticBag) As LocalSymbol
        '    Return Me.m_containingBinder.DeclareImplicitLocalVariable(nameSyntax, diagnostics)
        'End Function

    End Class

End Namespace


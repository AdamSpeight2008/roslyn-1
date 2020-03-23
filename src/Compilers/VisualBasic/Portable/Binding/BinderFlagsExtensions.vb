' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.
Imports System.Runtime.CompilerServices

Namespace Microsoft.CodeAnalysis.VisualBasic
    ''' <summary>
    ''' Extension methods for the <see cref="BinderFlags"/> type.
    ''' </summary>
    Friend Module BinderFlagsExtensions

        <Extension>
        Public Function AddAndRemove(self As BinderFlags, addThese As BinderFlags, removeThese As BinderFlags) As BinderFlags
            Return self.SetFlags(addThese).ClearFlags(removeThese)
        End Function

        <Extension>
        Public Function SetFlags(self As BinderFlags, theseFlags As BinderFlags) As BinderFlags
            Return self Or theseFlags
        End Function

        <Extension>
        Public Function ClearFlags(self As BinderFlags, theseFlags As BinderFlags) As BinderFlags
            Return self And Not theseFlags
        End Function

        <Extension>
        Public Function Includes(self As BinderFlags, other As BinderFlags) As Boolean
            Return (self And other) = other
        End Function

        <Extension>
        Public Function IncludesAny(self As BinderFlags, other As BinderFlags) As Boolean
            Return (self And other) <> 0
        End Function
        ''' <summary>
        ''' In case Flags is not set we want to default to Check
        ''' </summary>
        ''' <param name="flags"></param>
        ''' <returns></returns>
        <Extension>
        Friend Function IsCheckedRegion(flags As BinderFlags) As Boolean
            Return If(flags.Includes(BinderFlags.UncheckedRegion), False, True)
        End Function

    End Module
End Namespace

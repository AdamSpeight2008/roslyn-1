' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Namespace Microsoft.CodeAnalysis.VisualBasic
    ''' <summary>
    ''' A specific location for binding.
    ''' </summary>
    <Flags>
    Friend Enum BinderFlags As UInteger
        None ' No specific location
        ''' <summary>Remarks, mutually exclusive with <see cref="UncheckedRegion"/>.</summary>
        CheckedRegion = 1 << 14
        ''' <summary>Remarks, mutually exclusive with <see cref="CheckedRegion"/>.</summary>
        UncheckedRegion = 1 << 15
    End Enum
End Namespace

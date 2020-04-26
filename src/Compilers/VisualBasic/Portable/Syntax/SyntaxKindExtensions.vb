' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Runtime.CompilerServices

Namespace Microsoft.CodeAnalysis.VisualBasic
    Friend Module SyntaxKindExtensions

        ''' <summary>
        ''' Determine if the given <see cref="SyntaxKind"/> array contains the given kind.
        ''' </summary>
        ''' <param name="kinds">Array to search</param>
        ''' <param name="kind">Sought value</param>
        ''' <returns>True if <paramref name="kinds"/> contains the value <paramref name="kind"/>.</returns>
        ''' <remarks>PERF: Not using Array.IndexOf here because it results in a call to IndexOf on the default EqualityComparer for SyntaxKind. The default comparer for SyntaxKind is
        ''' the ObjectEqualityComparer which results in boxing allocations.</remarks>
        <Extension()>
        Public Function Contains(kinds As SyntaxKind(), kind As SyntaxKind) As Boolean
            For Each k In kinds
                If k = kind Then Return True
            Next
            Return False
        End Function

#Region "IsKindEither Extension methods"

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind) As Boolean
            Return (kind = kind0) Or (kind = kind1)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1) Or (kind = kind2)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1) Or kind.IsKindEither(kind2,kind3)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2) Or kind.IsKindEither(kind3,kind4)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind, kind5 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2) Or kind.IsKindEither(kind3, kind4, kind5)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind, kind5 As SyntaxKind, kind6 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2, kind3) Or kind.IsKindEither(kind4, kind5, kind6)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind, kind5 As SyntaxKind, kind6 As SyntaxKind, kind7 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2, kind3) Or kind.IsKindEither(kind4, kind5, kind6, kind7)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind, kind5 As SyntaxKind, kind6 As SyntaxKind, kind7 As SyntaxKind,
                                                          kind8 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2, kind3, kind4) Or kind.IsKindEither(kind5, kind6, kind7, kind8)
        End Function

        <Extension>
        Friend Function IsKindEither(kind As SyntaxKind,  kind0 As SyntaxKind, kind1 As SyntaxKind, kind2 As SyntaxKind, kind3 As SyntaxKind,
                                                          kind4 As SyntaxKind, kind5 As SyntaxKind, kind6 As SyntaxKind, kind7 As SyntaxKind,
                                                          kind8 As SyntaxKind, kind9 As SyntaxKind) As Boolean
            Return kind.IsKindEither(kind0, kind1, kind2, kind3, kind4) Or kind.IsKindEither(kind5, kind6, kind7, kind8, kind9)
        End Function

#End Region

    End Module

End Namespace


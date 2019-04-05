' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports Microsoft.CodeAnalysis.VisualBasic.Syntax.InternalSyntax
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Features.LangaugeFeatureService
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Features
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Version

Namespace Microsoft.CodeAnalysis.VisualBasic.Language.Version
  Friend NotInheritable Class VisualBasicRequiredLanguageVersionService

    Private Shared ReadOnly Property _RequiredVersionsForFeature As New Dictionary(Of Feature, VisualBasicRequiredLanguageVersion)()
    Public Shared ReadOnly Instance As New VisualBasicRequiredLanguageVersionService

     Shared Sub New()
      For Each f As Feature In LangaugeFeatureService.Instance.EnumerateLanguageFeatures
        Dim requiredVersion = New VisualBasicRequiredLanguageVersion(LangaugeFeatureService.Instance.GetLanguageVersion(f))
        _RequiredVersionsForFeature.Add(f, requiredVersion)
      Next
    End Sub

    Public Function GetRequiredLanguageVersion(f As Feature) As VisualBasicRequiredLanguageVersion
      Return _RequiredVersionsForFeature(f)
    End Function

  End Class

  Friend Class VisualBasicRequiredLanguageVersion
    Inherits RequiredLanguageVersion

    Friend ReadOnly Property Version As LanguageVersionService.LanguageVersion

    Friend Sub New(version As LanguageVersionService.LanguageVersion)
      Me.Version = version
    End Sub

    Public Overrides Function ToString() As String
      Return LanguageVersionService.Instance.ToDisplayString(Version)
    End Function

  End Class

End Namespace

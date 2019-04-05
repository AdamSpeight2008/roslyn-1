' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.
Imports Microsoft.CodeAnalysis.VisualBasic.Language.Version

Public Class LanguageVersionTests
    <Fact>
    Public Sub CurrentVersion()
        Dim highest = LanguageVersionService.Instance.EnumerateLanguageVersions.Where(Function(x) x <> LanguageVersionService.LanguageVersion.Latest).Max()

        Assert.Equal(LanguageVersionService.Instance.CurrentVersion, highest)
    End Sub
End Class

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Namespace Microsoft.CodeAnalysis.VisualBasic

    Friend NotInheritable Class VBDiagnostic
        Inherits DiagnosticWithInfo

        Friend Sub New(info As DiagnosticInfo, location As Location, Optional isSuppressed As Boolean = False)
            MyBase.New(info, location, isSuppressed)
        End Sub

        Public Overrides Function ToString() As String
            Return VisualBasicDiagnosticFormatter.Instance.Format(Me)
        End Function

        Friend Overrides Function WithLocation(location As Location) As Diagnostic
            If location Is Nothing Then Throw New ArgumentNullException(NameOf(location))
            If location IsNot Me.Location Then Return New VBDiagnostic(Info, location, IsSuppressed)
            Return Me
        End Function

        Friend Overrides Function WithSeverity(severity As DiagnosticSeverity) As Diagnostic
           Return If( Me.Severity = severity, Me, New VBDiagnostic(Info.GetInstanceWithSeverity(severity), Location, IsSuppressed))
        End Function

        Friend Overrides Function WithIsSuppressed(isSuppressed As Boolean) As Diagnostic
            Return If( Me.IsSuppressed = isSuppressed, Me, New VBDiagnostic(Info, Location, isSuppressed))
        End Function

    End Class

End Namespace


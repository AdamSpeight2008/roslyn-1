Imports Microsoft.CodeAnalysis.VisualBasic
Imports System.Runtime.CompilerServices
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Global.Microsoft.CodeAnalysis.VisualBasic
  Friend Module Exts
    <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind) As Boolean
        Return k = k0 OrElse k = k1
    End Function

  End Module
End Namespace

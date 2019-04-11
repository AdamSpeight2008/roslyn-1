Imports Microsoft.CodeAnalysis.VisualBasic
Imports System.Runtime.CompilerServices
Imports CoreInternalSyntax = Microsoft.CodeAnalysis.Syntax.InternalSyntax

Namespace Global.Microsoft.CodeAnalysis.VisualBasic

  Friend Module Exts

    <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind) As Boolean
        Return k = k0 OrElse k = k1
    End Function
    <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind, k2 As SyntaxKind) As Boolean
        Return k.IsEither(k0,k1) OrElse k=k2
    End Function
            <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind, k2 As SyntaxKind, k3 As SyntaxKind) As Boolean
        Return k.IsEither(k0, k1) OrElse k.IsEither(k2,k3)
    End Function
    <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind, k2 As SyntaxKind, k3 As SyntaxKind, k4 As SyntaxKind) As Boolean
        Return k.IsEither(k0,k1) OrElse k.IsEither(k2,k3) OrElse k=k4
    End Function
            <Extension>
    Friend Function IsEither(k As SyntaxKind, k0 As SyntaxKind, k1 As SyntaxKind, k2 As SyntaxKind, k3 As SyntaxKind, k4 As SyntaxKind, k5 As SyntaxKind) As Boolean
        Return k.IsEither(k0, k1) OrElse k.IsEither(k2,k3) OrElse k.IsEither(k4,k5)
    End Function

    <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char) As Boolean
        Return ch = ch0 OrElse ch = ch1
    End Function
    <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char) As Boolean
        Return ch = ch0 OrElse ch = ch1 OrElse ch = ch2
    End Function
     <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char, ch3 As Char) As Boolean
        Return ch.IsEither(ch0, ch1) OrElse ch.IsEither(ch2, ch3)
    End Function

     <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char, ch3 As Char, ch4 As Char) As Boolean
        Return ch.IsEither(ch0, ch1) OrElse ch.IsEither(ch2, ch3) OrElse ch = ch4
    End Function

    <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char, ch3 As Char, ch4 As Char, ch5 As Char) As Boolean
        Return ch.IsEither(ch0, ch1) OrElse ch.IsEither(ch2, ch3) OrElse ch.IsEither(ch4, ch5)
    End Function

    <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char, ch3 As Char, ch4 As Char, ch5 As Char, ch6 As Char) As Boolean
      Return ch.IsEither(ch0, ch1) OrElse ch.IsEither(ch2, ch3) OrElse ch.IsEither(ch4, ch5) OrElse ch = ch6
    End Function

    <Extension>
    Friend Function IsEither(ch As Char, ch0 As Char, ch1 As Char, ch2 As Char, ch3 As Char, ch4 As Char, ch5 As Char, ch6 As Char, ch7 As Char) As Boolean
      Return ch.IsEither(ch0, ch1) OrElse ch.IsEither(ch2, ch3) OrElse ch.IsEither(ch4, ch5) OrElse ch.IsEither(ch6, ch7)
    End Function

    <Extension>
    Friend Function IsBetween(value As Int32, Vmin AS Int32, VMax As Int32) As Boolean
      Return (Vmin <= Value) And (Value <= Vmax)
    End Function

  End Module

End Namespace

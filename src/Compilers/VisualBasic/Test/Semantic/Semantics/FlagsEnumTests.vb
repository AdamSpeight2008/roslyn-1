' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.SpecialType
Imports Microsoft.CodeAnalysis.Test.Utilities
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.OverloadResolution
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Microsoft.CodeAnalysis.VisualBasic.UnitTests.Emit

Namespace Microsoft.CodeAnalysis.VisualBasic.UnitTests.Semantics

    Public Class FlagsEnum
        Inherits BasicTestBase
        <Fact>
        Public Sub Test1()
            Dim p As new VisualBasicParseOptions(LanguageVersion.Latest)

            Dim compilationDef =
<compilation name="FlagsEnumOperations">
    <file name="a.vb">
Option Strict On

Module Module1

    &lt;System.Flags()>
    Enum X
       A = 1 &lt;&lt; 0
       B = 1 &lt;&lt; 1
       C = 1 &lt;&lt; 2
       D = 1 &lt;&lt; 3
    End Enum
            
    Public Sub Main()
       Dim s As X = X.A
       System.Console.WriteLine(Check0(s))
       System.Console.WriteLine(Check1(s))
    End Sub
    Private Function Check0(s As X) As Boolean
       Return s!A
    End Function
    Private Function Check1(s As X) As Boolean
       Return s!B
    End Function

End Module
    </file>
</compilation>

         Dim cv=   CompileAndVerify(compilationDef,
                             expectedOutput:=<![CDATA[
True
False
]]>,
parseOptions:= p)
            cv.VerifyIL("Module1.Check0",
                        "{
  // Code size        7 (0x7)
  .maxstack  2
  IL_0000:  ldarg.0
  IL_0001:  ldc.i4.1
  IL_0002:  and
  IL_0003:  ldc.i4.1
  IL_0004:  ceq
  IL_0006:  ret
}
")
            End Sub
    End Class

End Namespace


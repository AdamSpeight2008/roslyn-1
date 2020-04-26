' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.
Imports System.Runtime.InteropServices
Imports Microsoft.CodeAnalysis.PooledObjects
Imports Microsoft.CodeAnalysis.VisualBasic.Symbols

Namespace Microsoft.CodeAnalysis.VisualBasic

    Partial Friend Class Binder

        ''' <summary>
        ''' Given query operator source, infer control variable type from available
        ''' 'Select' methods. 
        ''' 
        ''' Returns inferred type or Nothing.
        ''' </summary>
        Private Function InferControlVariableType(
                                                   source As BoundExpression,
                                                   diagnostics As DiagnosticBag
                                                 ) As TypeSymbol
            Debug.Assert(source.IsValue)

            Dim result As TypeSymbol = Nothing

            ' Look for Select methods available for the source.
            Dim lookupResult As LookupResult = LookupResult.GetInstance()
            Dim useSiteDiagnostics As HashSet(Of DiagnosticInfo) = Nothing
            LookupMember(lookupResult, source.Type, StringConstants.SelectMethod, 0, QueryOperatorLookupOptions, useSiteDiagnostics)

            If lookupResult.IsGood Then

                Dim failedDueToAnAmbiguity As Boolean = False

                ' Name lookup does not look for extension methods if it found a suitable
                ' instance method, which is a good thing because according to language spec:
                '
                ' 11.21.2 Queryable Types
                ' ... , when determining the element type of a collection if there
                ' are instance methods that match well-known methods, then any extension methods
                ' that match well-known methods are ignored.
                Debug.Assert((QueryOperatorLookupOptions And LookupOptions.EagerlyLookupExtensionMethods) = 0)

                result = InferControlVariableType(lookupResult.Symbols, failedDueToAnAmbiguity)

                If result Is Nothing AndAlso Not failedDueToAnAmbiguity AndAlso Not lookupResult.Symbols(0).IsReducedExtensionMethod() Then
                    ' We tried to infer from instance methods and there were no suitable 'Select' method, 
                    ' let's try to infer from extension methods.
                    lookupResult.Clear()
                    Me.LookupExtensionMethods(lookupResult, source.Type, StringConstants.SelectMethod, 0, QueryOperatorLookupOptions, useSiteDiagnostics)

                    If lookupResult.IsGood Then
                        result = InferControlVariableType(lookupResult.Symbols, failedDueToAnAmbiguity)
                    End If
                End If
            End If

            diagnostics.Add(source, useSiteDiagnostics)
            lookupResult.Free()

            Return result
        End Function

        ''' <summary>
        ''' Given a set of 'Select' methods, infer control variable type. 
        ''' 
        ''' Returns inferred type or Nothing.
        ''' </summary>
        Private Function InferControlVariableType(
                                                   methods As ArrayBuilder(Of Symbol),
                                       <Out> ByRef failedDueToAnAmbiguity As Boolean
                                                 ) As TypeSymbol
            Dim result As TypeSymbol = Nothing
            failedDueToAnAmbiguity = False

            For Each method As MethodSymbol In methods
                Dim inferredType As TypeSymbol = InferControlVariableType(method)

                If inferredType IsNot Nothing Then
                    If inferredType.ReferencesMethodsTypeParameter(method) Then
                        failedDueToAnAmbiguity = True
                        Return Nothing
                    End If

                    If result Is Nothing Then
                        result = inferredType
                    ElseIf Not result.IsSameTypeIgnoringAll(inferredType) Then
                        failedDueToAnAmbiguity = True
                        Return Nothing
                    End If
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Given a method, infer control variable type. 
        ''' 
        ''' Returns inferred type or Nothing.
        ''' </summary>
        Private Function InferControlVariableType(
                                                   method As MethodSymbol
                                                 ) As TypeSymbol
            ' Ignore Subs
            If method.IsSub Then Return Nothing

            ' Only methods taking exactly one parameter are acceptable.
            If method.ParameterCount <> 1 Then Return Nothing

            Dim selectParameter As ParameterSymbol = method.Parameters(0)

            If selectParameter.IsByRef Then Return Nothing

            Dim parameterType As TypeSymbol = selectParameter.Type

            ' We are expecting a delegate type with the following shape:
            '     Function Selector (element as ControlVariableType) As AType

            ' The delegate type, directly converted to or argument of Expression(Of T)
            Dim delegateType As NamedTypeSymbol = parameterType.DelegateOrExpressionDelegate(Me)

            If delegateType Is Nothing Then Return Nothing

            Dim invoke As MethodSymbol = delegateType.DelegateInvokeMethod

            If invoke Is Nothing OrElse
               invoke.IsSub OrElse
               invoke.ParameterCount <> 1 Then Return Nothing

            Dim invokeParameter As ParameterSymbol = invoke.Parameters(0)

            ' Do not allow Optional, ParamArray and ByRef.
            If invokeParameter.IsOptional OrElse
               invokeParameter.IsByRef OrElse
               invokeParameter.IsParamArray Then Return Nothing

            Dim controlVariableType As TypeSymbol = invokeParameter.Type

            Return If(controlVariableType.IsErrorType(), Nothing, controlVariableType)
        End Function

    End Class

End Namespace

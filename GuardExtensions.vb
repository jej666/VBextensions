Imports System.Diagnostics
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Globalization

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module GuardExtensions

    ''' <summary>
    ''' Asserts the not null.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <param name="selector">The selector.</param>
    ''' <exception cref="System.ArgumentNullException"></exception>
    <DebuggerStepThrough(), Extension()> _
    Public Sub AssertNotNull(Of T)(argumentValue As T, selector As Expression(Of Func(Of T)))
        If argumentValue Is Nothing Then
            Dim name As String = DirectCast(selector.Body, MemberExpression).Member.Name
            Throw New ArgumentNullException(name)
        End If
    End Sub

    ''' <summary>
    ''' Asserts the not empty.
    ''' </summary>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <param name="selector">The selector.</param>
    ''' <exception cref="System.ArgumentNullException"></exception>
    ''' <exception cref="System.ArgumentException">String must not be empty.</exception>
    <DebuggerStepThrough(), Extension()> _
    Public Sub AssertNotEmpty(argumentValue As String, selector As Expression(Of Func(Of String)))
        If argumentValue Is Nothing Then
            Dim paramName As String = DirectCast(selector.Body, MemberExpression).Member.Name
            Throw New ArgumentNullException(paramName)
        End If
        If String.IsNullOrEmpty(argumentValue) Then
            Dim paramName As String = DirectCast(selector.Body, MemberExpression).Member.Name
            Throw New ArgumentException("String must not be empty.", paramName)
        End If
    End Sub

    ''' <summary>
    ''' Throws the exception if negative or null.
    ''' </summary>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough(), Extension()>
    Public Sub AssertNotNegativeOrNull(argumentValue As Integer)
        If argumentValue <= 0 Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Value {0} cannot be negative or null", argumentValue))
        End If
    End Sub

    ''' <summary>
    ''' Throws the exception if negative or null.
    ''' </summary>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough(), Extension()>
    Public Sub AssertNotNegativeOrNull(argumentValue As Long)
        If argumentValue <= 0 Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Value {0} cannot be negative or null", argumentValue))
        End If
    End Sub

    ''' <summary>
    ''' Throws the exception if negative or null.
    ''' </summary>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough(), Extension()>
    Public Sub AssertNotNegativeOrNull(argumentValue As Int16)
        If argumentValue <= 0 Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Value {0} cannot be negative or null", argumentValue))
        End If
    End Sub

    ''' <summary>
    ''' Throws the exception if negative or null.
    ''' </summary>
    ''' <param name="argumentValue">The argument value.</param>
    ''' <remarks></remarks>
    <DebuggerStepThrough(), Extension()>
    Public Sub AssertNotNegativeOrNull(argumentValue As Decimal)
        If argumentValue <= 0 Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Value {0} cannot be negative or null", argumentValue))
        End If
    End Sub

End Module

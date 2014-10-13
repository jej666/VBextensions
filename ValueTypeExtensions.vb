Imports System.Runtime.CompilerServices

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module ValueTypeExtensions

    ''' <summary>
    ''' 	Determines whether the specified value is empty.
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "value">The value.</param>
    ''' <returns>
    ''' 	<c>true</c> if the specified value is empty; otherwise, <c>false</c>.
    ''' </returns>
    <Extension()> _
    Public Function IsEmpty(Of T As Structure)(value As T) As Boolean
        Return value.Equals(Nothing)
    End Function

    ''' <summary>
    ''' 	Determines whether the specified value is not empty.
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "value">The value.</param>
    ''' <returns>
    ''' 	<c>true</c> if the specified value is not empty; otherwise, <c>false</c>.
    ''' </returns>
    <Extension()> _
    Public Function IsNotEmpty(Of T As Structure)(value As T) As Boolean
        Return (value.IsEmpty() = False)
    End Function

    ''' <summary>
    ''' 	Converts the specified value to a corresponding nullable type
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "value">The value.</param>
    ''' <returns>The nullable type</returns>
    <Extension()> _
    Public Function ToNullable(Of T As Structure)(value As T) As Nullable(Of T)
        Return (If(value.IsEmpty(), Nothing, CType(value, Nullable(Of T))))
    End Function

    ''' <summary>
    ''' Determines whether the specified value is null.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="value">The value.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function IsNull(Of T As Structure)(value As T?) As Boolean
        Return Not value.HasValue
    End Function

End Module

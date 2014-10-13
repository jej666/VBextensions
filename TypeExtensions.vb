Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Globalization

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module TypeExtensions

    ''' <summary>
    ''' Checks whether or not the
    ''' </summary>
    ''' <param name="val">The value to test for conversion to the type
    ''' associated with the property</param>
    <Extension()> _
    Public Function CanConvertTo(Of T)(val As String) As Boolean
        Return GetType(T).CanConvertTo(val)
    End Function

    ''' <summary>
    ''' Checks whether or not the
    ''' </summary>
    ''' <param name="val">The value to test for conversion to the type
    ''' associated with the property</param>
    <Extension()> _
    Public Function CanConvertTo(type As Type, val As String) As Boolean
        ' Data could be passed as string value.
        ' Try to change type to check type safety.
        Try
            If type = GetType(Integer) Then
                If Integer.TryParse(val, 0) Then
                    Return True
                End If

                Return False
            ElseIf type = GetType(String) Then
                Return True
            ElseIf type = GetType(Double) Then
                If Double.TryParse(val, 0) Then
                    Return True
                End If

                Return False
            ElseIf type = GetType(Long) Then
                If Long.TryParse(val, 0) Then
                    Return True
                End If

                Return False
            ElseIf type = GetType(Single) Then
                If Single.TryParse(val, 0) Then
                    Return True
                End If

                Return False
            ElseIf type = GetType(Boolean) Then
                If Boolean.TryParse(val, False) Then
                    Return True
                End If

                Return False
            ElseIf type = GetType(DateTime) Then
                Dim d As DateTime = DateTime.MinValue
                If DateTime.TryParse(val, d) Then
                    Return True
                End If

                Return False
            ElseIf type.BaseType = GetType([Enum]) Then
                [Enum].Parse(type, val, True)
            End If
        Catch generatedExceptionName As Exception
            Return False
        End Try

        'Conversion worked.
        Return True
    End Function

    ''' <summary>
    ''' Checks whether or not the
    ''' </summary>
    ''' <param name="propInfo">The property represnting the type to convert
    ''' val to</param>
    ''' <param name="val">The value to test for conversion to the type
    ''' associated with the property</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function CanConvertToCorrectType(propInfo As PropertyInfo, val As String) As Boolean
        ' Data could be passed as string value.
        ' Try to change type to check type safety.
        Try
            If propInfo.PropertyType = GetType(Integer) Then
                If Integer.TryParse(val, 0) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType = GetType(String) Then
                Return True
            ElseIf propInfo.PropertyType = GetType(Double) Then
                If Double.TryParse(val, 0) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType = GetType(Long) Then
                If Long.TryParse(val, 0) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType = GetType(Single) Then
                If Single.TryParse(val, 0) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType = GetType(Boolean) Then
                If Boolean.TryParse(val, False) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType = GetType(DateTime) Then
                Dim d As DateTime = DateTime.MinValue
                If DateTime.TryParse(val, d) Then
                    Return True
                End If
                Return False
            ElseIf propInfo.PropertyType.BaseType = GetType([Enum]) Then
                [Enum].Parse(propInfo.PropertyType, val, True)
            End If
        Catch generatedExceptionName As Exception
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Convert the val from string type to the same time as the property.
    ''' </summary>
    ''' <param name="propInfo">Property representing the type to convert to</param>
    ''' <param name="val">val to convert</param>
    ''' <returns>converted value with the same time as the property</returns>
    <Extension()> _
    Public Function ConvertToSameType(propInfo As PropertyInfo, val As Object) As Object
        Dim convertedType As Object = Nothing

        If propInfo.PropertyType = GetType(Integer) Then
            convertedType = Convert.ChangeType(val, GetType(Integer), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(Double) Then
            convertedType = Convert.ChangeType(val, GetType(Double), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(Long) Then
            convertedType = Convert.ChangeType(val, GetType(Long), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(Single) Then
            convertedType = Convert.ChangeType(val, GetType(Single), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(Boolean) Then
            convertedType = Convert.ChangeType(val, GetType(Boolean), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(DateTime) Then
            convertedType = Convert.ChangeType(val, GetType(DateTime), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType = GetType(String) Then
            convertedType = Convert.ChangeType(val, GetType(String), CultureInfo.CurrentCulture)
        ElseIf propInfo.PropertyType.BaseType = GetType([Enum]) AndAlso TypeOf val Is String Then
            convertedType = [Enum].Parse(propInfo.PropertyType, CStr(val), True)
        End If
        Return convertedType
    End Function

    ''' <summary>
    ''' Determine if the type of the property and the val are the same
    ''' </summary>
    ''' <param name="propInfo"></param>
    ''' <param name="val"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function IsSameType(propInfo As PropertyInfo, val As Object) As Boolean
        ' Quick Validation.
        If propInfo.PropertyType = GetType(Integer) AndAlso TypeOf val Is Integer Then
            Return True
        End If
        If propInfo.PropertyType = GetType(Boolean) AndAlso TypeOf val Is Boolean Then
            Return True
        End If
        If propInfo.PropertyType = GetType(String) AndAlso TypeOf val Is String Then
            Return True
        End If
        If propInfo.PropertyType = GetType(Double) AndAlso TypeOf val Is Double Then
            Return True
        End If
        If propInfo.PropertyType = GetType(Long) AndAlso TypeOf val Is Long Then
            Return True
        End If
        If propInfo.PropertyType = GetType(Single) AndAlso TypeOf val Is Single Then
            Return True
        End If
        If propInfo.PropertyType = GetType(DateTime) AndAlso TypeOf val Is DateTime Then
            Return True
        End If
        If TypeOf propInfo.PropertyType Is Object AndAlso propInfo.PropertyType.[GetType]() = val.[GetType]() Then
            Return True
        End If

        Return False
    End Function

End Module
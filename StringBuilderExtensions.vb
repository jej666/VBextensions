Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Globalization

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module StringBuilderExtensions

    ''' <summary>
    ''' AppendLine version with format string parameters.
    ''' </summary>
    <Extension()> _
    Public Sub AppendLine(builder As StringBuilder, value As String, ParamArray parameters As [Object]())
        builder.AppendLine(String.Format(CultureInfo.CurrentCulture, value, parameters))
    End Sub

    ''' <summary>
    ''' Appends the value of the object's System.Object.ToString() method followed by the default line terminator to the end of the current
    ''' System.Text.StringBuilder object if a condition is true
    ''' </summary>
    ''' <param name="builder"></param>
    ''' <param name="condition">The conditional expression to evaluate.</param>
    ''' <param name="value"></param>
    <Extension()> _
    Public Function AppendLineIf(builder As StringBuilder, condition As Boolean, value As Object) As StringBuilder
        If condition Then
            builder.AppendLine(value.ToString())
        End If
        Return builder
    End Function

    ''' <summary>
    ''' Appends the string returned by processing a composite format string, which contains zero or more format items, followed by the default
    ''' line terminator to the end of the current System.Text.StringBuilder object if a condition is true
    ''' </summary>
    ''' <param name="builder"></param>
    ''' <param name="condition">The conditional expression to evaluate.</param>
    ''' <param name="format"></param>
    ''' <param name="args"></param>
    <Extension()> _
    Public Function AppendLineIf(builder As StringBuilder, condition As Boolean, format As String, ParamArray args As Object()) As StringBuilder
        If condition Then
            builder.AppendFormat(CultureInfo.CurrentCulture, format, args).AppendLine()
        End If
        Return builder
    End Function

    ''' <summary>
    ''' Appends the value of the object's System.Object.ToString() method to the end of the current
    ''' System.Text.StringBuilder object if a condition is true
    ''' </summary>
    ''' <param name="builder"></param>
    ''' <param name="condition"></param>
    ''' <param name="value"></param>
    <Extension()> _
    Public Function AppendIf(builder As StringBuilder, condition As Boolean, value As Object) As StringBuilder
        If condition Then
            builder.Append(value.ToString())
        End If
        Return builder
    End Function

    ''' <summary>
    ''' Appends the string returned by processing a composite format string, which contains zero or more format items,
    ''' to the end of the current System.Text.StringBuilder object if a condition is true
    ''' </summary>
    ''' <param name="builder"></param>
    ''' <param name="condition"></param>
    ''' <param name="format"></param>
    ''' <param name="args"></param>
    <Extension()> _
    Public Function AppendFormatIf(builder As StringBuilder, condition As Boolean, format As String, ParamArray args As Object()) As StringBuilder
        If condition Then
            builder.AppendFormat(CultureInfo.CurrentCulture, format, args)
        End If
        Return builder
    End Function

End Module

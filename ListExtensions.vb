Imports System.Runtime.CompilerServices
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Text
Imports System.Collections.Generic

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module ListExtensions

    ''' <summary>
    ''' Returns a delimited string from the list.
    ''' </summary>
    ''' <param name="ls"></param>
    ''' <param name="delimiter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ToDelimitedString(ls As List(Of String), delimiter As String) As String
        Dim sb As New StringBuilder
        For Each buf As String In ls
            sb.Append(buf)
            sb.Append(delimiter)
        Next
        Return sb.ToString.Trim(CChar(delimiter))
    End Function

    ''' <summary>
    ''' 	Return the index of the first matching item or -1.
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "list">The list.</param>
    ''' <param name = "comparison">The comparison.</param>
    ''' <returns>The item index</returns>
    <Extension()> _
    Public Function IndexOf(Of T)(list As IList(Of T), comparison As Func(Of T, Boolean)) As Integer
        For i As Integer = 0 To list.Count - 1
            If comparison(list(i)) Then
                Return i
            End If
        Next
        Return -1
    End Function

    

End Module

Imports System.Runtime.CompilerServices

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module NumericValueExtensions

    ''' <summary>
    ''' 	Performs the specified action n times based on the underlying int value.
    ''' </summary>
    ''' <param name = "value">The value.</param>
    ''' <param name = "action">The action.</param>
    <Extension()> _
    Public Sub ForEachTimes(value As Integer, action As Action)
        value.ForEachTimes(action)
    End Sub

    ''' <summary>
    ''' 	Performs the specified action n times based on the underlying int value.
    ''' </summary>
    ''' <param name = "value">The value.</param>
    ''' <param name = "action">The action.</param>
    <Extension()> _
    Public Sub ForEachTimes(value As Integer, action As Action(Of Integer))
        For i As Integer = 0 To value - 1
            action(i)
        Next
    End Sub

    ''' <summary>
    ''' 	Performs the specified action n times based on the underlying int value.
    ''' </summary>
    ''' <param name = "value">The value.</param>
    ''' <param name = "startIndex">The start index.</param>
    ''' <param name = "action">The action.</param>
    <Extension()> _
    Public Sub ForEachTimes(value As Integer, startIndex As Integer, action As Action(Of Integer))
        For i As Integer = startIndex To value - 1
            action(i)
        Next
    End Sub

    ''' <summary>
    ''' 	Performs the specified action n times based on the underlying int value.
    ''' </summary>
    ''' <param name = "value">The value.</param>
    ''' <param name = "startIndex">The start index.</param>
    ''' <param name = "predicate">The action.</param>
    <Extension()> _
    Public Function ForEachTimes(Of TResult)(value As Integer, startIndex As Integer, predicate As Func(Of Integer, TResult)) As TResult
        For i As Integer = startIndex To value - 1
            Return predicate(i)
        Next
    End Function

End Module

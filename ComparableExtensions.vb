Imports System.Runtime.CompilerServices
Imports System.Collections.Generic

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module ComparableExtensions

    ''' <summary>
    ''' 	Determines whether the specified value is between the the defined minimum and maximum range (including those values).
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "value">The value.</param>
    ''' <param name = "minValue">The minimum value.</param>
    ''' <param name = "maxValue">The maximum value.</param>
    ''' <returns>
    ''' 	<c>true</c> if the specified value is between min and max; otherwise, <c>false</c>.
    ''' </returns>
    ''' <example>
    ''' 	var value = 5;
    ''' 	if(value.IsBetween(1, 10)) {
    ''' 	// ...
    ''' 	}
    ''' </example>
    <Extension()> _
    Public Function IsBetween(Of T As IComparable(Of T))(value As T, minValue As T, maxValue As T) As Boolean
        Return value.IsBetween(minValue, maxValue, Comparer(Of T).[Default])
    End Function

    ''' <summary>
    ''' 	Determines whether the specified value is between the the defined minimum and maximum range (including those values).
    ''' </summary>
    ''' <typeparam name = "T"></typeparam>
    ''' <param name = "value">The value.</param>
    ''' <param name = "minValue">The minimum value.</param>
    ''' <param name = "maxValue">The maximum value.</param>
    ''' <param name = "comparer">An optional comparer to be used instead of the types default comparer.</param>
    ''' <returns>
    ''' 	<c>true</c> if the specified value is between min and max; otherwise, <c>false</c>.
    ''' </returns>
    ''' <example>
    ''' 	var value = 5;
    ''' 	if(value.IsBetween(1, 10)) {
    ''' 	// ...
    ''' 	}
    ''' </example>
    <Extension()> _
    Public Function IsBetween(Of T As IComparable(Of T))(value As T, minValue As T, maxValue As T, comparer As IComparer(Of T)) As Boolean
        comparer.AssertNotNull(Function() comparer)
        Dim minMaxCompare = comparer.Compare(minValue, maxValue)
        Return If(minMaxCompare < 0, ((comparer.Compare(value, minValue) >= 0) AndAlso (comparer.Compare(value, maxValue) <= 0)),
                  ((comparer.Compare(value, maxValue) >= 0) AndAlso (comparer.Compare(value, minValue) <= 0)))
    End Function

End Module

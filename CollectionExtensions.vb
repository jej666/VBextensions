Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices

''' <summary>
''' 	Extension methods for all kind of Collections implementing the ICollection&lt;T&gt; interface
''' </summary>
<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module CollectionExtensions

    ''' <summary>
    ''' 	Adds a value uniquely to a collection and returns a value whether the value was added or not.
    ''' </summary>
    ''' <typeparam name = "T">The generic collection value type</typeparam>
    ''' <param name = "collection">The collection.</param>
    ''' <param name = "value">The value to be added.</param>
    ''' <returns>Indicates whether the value was added or not</returns>
    ''' <example>
    ''' 	<code>
    ''' 		list.AddUnique(1); // returns true;
    ''' 		list.AddUnique(1); // returns false the second time;
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function AddUnique(Of T)(collection As ICollection(Of T), value As T) As Boolean
        Dim alreadyHas = collection.Contains(value)
        If Not alreadyHas Then
            collection.Add(value)
        End If
        Return alreadyHas
    End Function

    ''' <summary>
    ''' Adds the range.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="collection">The collection.</param>
    ''' <param name="values">The values.</param>
    ''' <remarks></remarks>
    <Extension()> _
    Public Sub AddRange(Of T)(collection As ICollection(Of T), values As IEnumerable(Of T))
        collection.AssertNotNull(Function() collection)
        values.AssertNotNull(Function() values)
        values.ForEach(Sub(value) collection.Add(value))
    End Sub

    ''' <summary>
    ''' 	Adds a range of value uniquely to a collection and returns the amount of values added.
    ''' </summary>
    ''' <typeparam name = "T">The generic collection value type.</typeparam>
    ''' <param name = "collection">The collection.</param>
    ''' <param name = "values">The values to be added.</param>
    ''' <returns>The amount if values that were added.</returns>
    <Extension()> _
    Public Function AddRangeUnique(Of T)(collection As ICollection(Of T), values As IEnumerable(Of T)) As Integer
        Return values.Count(Function(value) collection.AddUnique(value))
    End Function

    '''<summary>
    '''	Remove an item from the collection with predicate
    '''</summary>
    '''<param name = "collection"></param>
    '''<param name = "predicate"></param>
    '''<typeparam name = "T"></typeparam>
    '''<exception cref = "ArgumentNullException"></exception>
    <Extension()> _
    Public Sub RemoveWhere(Of T)(collection As ICollection(Of T), predicate As Predicate(Of T))
        collection.AssertNotNull(Function() collection)
        Dim deleteList = collection.Where(Function(child) predicate(child)).ToList()
        deleteList.ForEach(Function(at) collection.Remove(at))
    End Sub

    ''' <summary>
    ''' Tests if the collection is empty.
    ''' </summary>
    ''' <param name="collection">The collection to test.</param>
    ''' <returns>True if the collection is empty.</returns>
    <Extension()> _
    Public Function IsEmpty(collection As ICollection) As Boolean
        collection.AssertNotNull(Function() collection)
        Return collection.Count = 0
    End Function

    ''' <summary>
    ''' Tests if the collection is empty.
    ''' </summary>
    ''' <typeparam name="T">The type of the items in
    ''' the collection.</typeparam>
    ''' <param name="collection">The collection to test.</param>
    ''' <returns>True if the collection is empty.</returns>
    <Extension()> _
    Public Function IsEmpty(Of T)(collection As ICollection(Of T)) As Boolean
        collection.AssertNotNull(Function() collection)
        Return collection.Count = 0
    End Function

    ''' <summary>
    ''' Tests if the collection is empty.
    ''' </summary>
    ''' <param name="collection">The collection to test.</param>
    ''' <returns>True if the collection is empty.</returns>
    <Extension()> _
    Public Function IsEmpty(collection As IList) As Boolean
        collection.AssertNotNull(Function() collection)
        Return collection.Count = 0
    End Function

    ''' <summary>
    ''' Tests if the collection is empty.
    ''' </summary>
    ''' <typeparam name="T">The type of the items in
    ''' the collection.</typeparam>
    ''' <param name="collection">The collection to test.</param>
    ''' <returns>True if the collection is empty.</returns>
    <Extension()> _
    Public Function IsEmpty(Of T)(collection As IList(Of T)) As Boolean
        collection.AssertNotNull(Function() collection)
        Return collection.Count = 0
    End Function

    ''' <summary>
    ''' Teste si index compris entre 0 et GetUpperBound(0)
    ''' </summary>
    ''' <param name="collection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function IsInBounds(Of T)(collection As ICollection(Of T), value As Integer) As Boolean
        collection.AssertNotNull(Function() collection)
        Return value >= 0 AndAlso value < collection.Count
    End Function

    ''' <summary>
    ''' Teste si index compris entre 0 et GetUpperBound(0)
    ''' </summary>
    ''' <param name="collection"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function IsInBounds(Of T)(collection As IList(Of T), value As Integer) As Boolean
        collection.AssertNotNull(Function() collection)
        Return value >= 0 AndAlso value < collection.Count
    End Function

    ''' <summary>
    ''' Automatics the dictionary.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="items">The items.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function ToDictionary(Of T)(items As IList(Of T)) As IDictionary(Of T, T)
        Dim dict As IDictionary(Of T, T) = New Dictionary(Of T, T)()
        For Each item As T In items
            dict(item) = item
        Next
        Return dict
    End Function

End Module

Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.ComponentModel

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module EnumerableExtensions

    <Extension()>
    Public Function Count(Of T)(items As IEnumerable(Of T)) As Integer
        Dim c = TryCast(items, ICollection(Of T))
        If c IsNot Nothing Then
            Return c.Count
        End If
        Dim _count As Integer = 0
        ForEach(Of T)(items, Function(ignore) Math.Max(System.Threading.Interlocked.Increment(_count), _count - 1))
        Return _count
    End Function

    ''' <summary>
    ''' Determines whether an IEnumerable contains any item
    ''' </summary>
    ''' <param name="enumerable">the IEnumerable</param>
    ''' <returns>false if enumerable is null or contains no items</returns>
    <Extension()>
    Public Function HasItems(enumerable As IEnumerable) As Boolean
        If enumerable Is Nothing Then
            Return False
        End If
        Dim enumerator = enumerable.GetEnumerator()
        If enumerator IsNot Nothing AndAlso enumerator.MoveNext() Then
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Check if any of the items in the collection satisfied by the condition.
    ''' </summary>
    ''' <param name="items"></param>
    ''' <param name="executor"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function IsTrueForAny(Of T)(items As IEnumerable(Of T), executor As Func(Of T, Boolean)) As Boolean
        Return items.Any(Function(item) executor(item))
    End Function

    '''<summary>
    '''	Get Distinct
    '''</summary>
    '''<param name = "source"></param>
    '''<param name = "expression"></param>
    '''<typeparam name = "T"></typeparam>
    '''<typeparam name = "TKey"></typeparam>
    '''<returns></returns>
    <Extension()> _
    Public Function Distinct(Of T, TKey)(source As IEnumerable(Of T), expression As Func(Of T, TKey)) As IEnumerable(Of T)
        Return If(source Is Nothing, Enumerable.Empty(Of T)(), source.GroupBy(expression).Select(Function(i) i.First()))
    End Function

    ''' <summary>
    ''' Check for any nulls.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="items"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function HasAnyNulls(Of T)(items As IEnumerable(Of T)) As Boolean
        Return IsTrueForAny(Of T)(items, Function(x) x Is Nothing)
    End Function

    ''' <summary>
    ''' Converts an enumerable collection to an delimited string.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="items"></param>
    ''' <param name="delimiter"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function AsDelimited(Of T)(items As IEnumerable(Of T), delimiter As String) As String
        Dim itemList As New List(Of String)()
        For Each item As T In items
            itemList.Add(item.ToString())
        Next
        Return [String].Join(delimiter, itemList.ToArray())
    End Function

    ''' <summary>
    ''' Ases the binding list.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="items">The items.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function AsBindingList(Of T)(items As IEnumerable(Of T)) As BindingList(Of T)
        Dim itemList As New BindingList(Of T)()
        For Each item As T In items
            itemList.Add(item)
        Next
        Return itemList
    End Function

    ''' <summary>
    ''' Filters sequence that are between <paramref name="min"/> and <paramref name="max"/>.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the elements of source.</typeparam>
    ''' <param name="source">An <see cref="IEnumerable(Of TSource)"/> to filter.</param>
    ''' <param name="min">The minimum value.</param>
    ''' <param name="max">The maximum value.</param>
    ''' <returns>An <see cref="IEnumerable(of TSource)"/> that contains elements from the input sequence that satisfy the condition.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    <Extension()> _
    Public Function Between(Of TSource)(source As IEnumerable(Of TSource), min As TSource, max As TSource) As IEnumerable(Of TSource)
        Return source.Between(min, max, Comparer(Of TSource).[Default])
    End Function

    ''' <summary>
    ''' Filters sequence that are between <paramref name="min"/> and <paramref name="max"/>.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the elements of source.</typeparam>
    ''' <param name="source">An <see cref="IEnumerable(Of TSource)"/> to filter.</param>
    ''' <param name="min">The minimum value.</param>
    ''' <param name="max">The maximum value.</param>
    ''' <param name="comparer">The <see cref="IComparer(Of TSource)"/> to compare values.</param>
    ''' <returns>An <see cref="IEnumerable(Of TSource)"/> that contains elements from the input sequence that satisfy the condition.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    ''' <exception cref="ArgumentNullException"><paramref name="comparer"/> is null.</exception>
    <Extension()> _
    Public Function Between(Of TSource)(source As IEnumerable(Of TSource), min As TSource, max As TSource, comparer As IComparer(Of TSource)) As IEnumerable(Of TSource)
        Return source.Between(Function(x) x, min, max, comparer)
    End Function

    ''' <summary>
    ''' Filters sequence that are between <paramref name="min"/> and <paramref name="max"/>.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the elements of source.</typeparam>
    ''' <typeparam name="TKey">The type of the key returned by keySelector.</typeparam>
    ''' <param name="source">An <see cref="IEnumerable(Of TSource)"/> to filter.</param>
    ''' <param name="keySelector">A function to extract the key for each element.</param>
    ''' <param name="min">The minimum value.</param>
    ''' <param name="max">The maximum value.</param>
    ''' <returns>An <see cref="IEnumerable(Of TSource)"/> that contains elements from the input sequence that satisfy the condition.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    ''' <exception cref="ArgumentNullException"><paramref name="keySelector"/> is null.</exception>
    <Extension()> _
    Public Function Between(Of TSource, TKey)(source As IEnumerable(Of TSource), keySelector As Func(Of TSource, TKey), min As TKey, max As TKey) As IEnumerable(Of TSource)
        Return source.Between(keySelector, min, max, Comparer(Of TKey).[Default])
    End Function

    ''' <summary>
    ''' Filters sequence that are between <paramref name="min"/> and <paramref name="max"/>.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the elements of source.</typeparam>
    ''' <typeparam name="TKey">The type of the key returned by keySelector.</typeparam>
    ''' <param name="source">An <see cref="IEnumerable(Of TSource)"/> to filter.</param>
    ''' <param name="selector">A function to extract the key for each element.</param>
    ''' <param name="min">The minimum value.</param>
    ''' <param name="max">The maximum value.</param>
    ''' <param name="comparer">The <see cref="IComparer(Of TSource)"/> to compare values.</param>
    ''' <returns>An <see cref="IEnumerable(Of TSource)"/> that contains elements from the input sequence that satisfy the condition.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    ''' <exception cref="ArgumentNullException"><paramref name="keySelector"/> is null.</exception>
    ''' <exception cref="ArgumentNullException"><paramref name="comparer"/> is null.</exception>
    <Extension()> _
    Public Function Between(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey), min As TKey, max As TKey, comparer As IComparer(Of TKey)) As IEnumerable(Of TSource)
        source.AssertNotNull(Function() source)
        selector.AssertNotNull(Function() selector)
        comparer.AssertNotNull(Function() comparer)
        Return source.Where(Function(x) comparer.Compare(selector(x), min) >= 0 AndAlso comparer.Compare(selector(x), max) <= 0)
    End Function

    ''' <summary>
    ''' 	Returns the maximum item based on a provided selector.
    ''' </summary>
    ''' <typeparam name = "TItem">The item type</typeparam>
    ''' <typeparam name = "TValue">The value item</typeparam>
    ''' <param name = "items">The items.</param>
    ''' <param name = "selector">The selector.</param>
    ''' <param name = "maxValue">The max value as output parameter.</param>
    ''' <returns>The maximum item</returns>
    ''' <example>
    ''' 	<code>
    ''' 		int age;
    ''' 		var oldestPerson = persons.MaxItem(p =&gt; p.Age, out age);
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function MaxItem(Of TItem As Class, TValue As IComparable)(items As IEnumerable(Of TItem), selector As Func(Of TItem, TValue), ByRef maxValue As TValue) As TItem
        Dim _maxItem As TItem = Nothing
        maxValue = Nothing
        For Each item In items
            If item Is Nothing Then
                Continue For
            End If
            Dim itemValue = selector(item)
            If (_maxItem IsNot Nothing) AndAlso
                (itemValue.CompareTo(maxValue) <= 0) Then
                Continue For
            End If
            maxValue = itemValue
            _maxItem = item
        Next
        Return _maxItem
    End Function

    ''' <summary>
    ''' 	Returns the maximum item based on a provided selector.
    ''' </summary>
    ''' <typeparam name = "TItem">The item type</typeparam>
    ''' <typeparam name = "TValue">The value item</typeparam>
    ''' <param name = "items">The items.</param>
    ''' <param name = "selector">The selector.</param>
    ''' <returns>The maximum item</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var oldestPerson = persons.MaxItem(p =&gt; p.Age);
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function MaxItem(Of TItem As Class, TValue As IComparable)(items As IEnumerable(Of TItem), selector As Func(Of TItem, TValue)) As TItem
        Dim maxValue As TValue
        Return items.MaxItem(selector, maxValue)
    End Function

    ''' <summary>
    ''' 	Returns the minimum item based on a provided selector.
    ''' </summary>
    ''' <typeparam name = "TItem">The item type</typeparam>
    ''' <typeparam name = "TValue">The value item</typeparam>
    ''' <param name = "items">The items.</param>
    ''' <param name = "selector">The selector.</param>
    ''' <param name = "minValue">The min value as output parameter.</param>
    ''' <returns>The minimum item</returns>
    ''' <example>
    ''' 	<code>
    ''' 		int age;
    ''' 		var youngestPerson = persons.MinItem(p =&gt; p.Age, out age);
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function MinItem(Of TItem As Class, TValue As IComparable)(items As IEnumerable(Of TItem), selector As Func(Of TItem, TValue), minValue As TValue) As TItem
        Dim minItem_ As TItem = Nothing
        minValue = Nothing
        For Each item In items
            If item Is Nothing Then
                Continue For
            End If
            Dim itemValue = selector(item)
            If (minItem_ IsNot Nothing) AndAlso
                (itemValue.CompareTo(minValue) >= 0) Then
                Continue For
            End If
            minValue = itemValue
            minItem_ = item
        Next
        Return minItem_
    End Function

    ''' <summary>
    ''' 	Returns the minimum item based on a provided selector.
    ''' </summary>
    ''' <typeparam name = "TItem">The item type</typeparam>
    ''' <typeparam name = "TValue">The value item</typeparam>
    ''' <param name = "items">The items.</param>
    ''' <param name = "selector">The selector.</param>
    ''' <returns>The minimum item</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var youngestPerson = persons.MinItem(p =&gt; p.Age);
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function MinItem(Of TItem As Class, TValue As IComparable)(items As IEnumerable(Of TItem), selector As Func(Of TItem, TValue)) As TItem
        Dim minValue As TValue
        Return items.MinItem(selector, minValue)
    End Function

    ''' <summary>
    ''' Returns the minimal element of the given sequence, based on
    ''' the given projection.
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the minimal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current minimal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <returns>The minimal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="selector"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>
    <Extension()> _
    Public Function MinBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey)) As TSource
        Return source.MinBy(selector, Comparer(Of TKey).[Default])
    End Function

    ''' <summary>
    ''' Returns the minimal element of the given sequence, based on
    ''' the given projection and the specified comparer for projected values.
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the minimal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current minimal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <param name="comparer">Comparer to use to compare projected values</param>
    ''' <returns>The minimal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/>, <paramref name="selector"/>
    ''' or <paramref name="comparer"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>
    <Extension()> _
    Public Function MinBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey), comparer As IComparer(Of TKey)) As TSource
        source.AssertNotNull(Function() source)
        selector.AssertNotNull(Function() selector)
        comparer.AssertNotNull(Function() comparer)
        Using sourceIterator = source.GetEnumerator()
            If Not sourceIterator.MoveNext() Then
                Throw New InvalidOperationException("Sequence contains no elements")
            End If
            Dim min = sourceIterator.Current
            Dim minKey = selector(min)
            While sourceIterator.MoveNext()
                Dim candidate = sourceIterator.Current
                Dim candidateProjected = selector(candidate)
                If comparer.Compare(candidateProjected, minKey) < 0 Then
                    min = candidate
                    minKey = candidateProjected
                End If
            End While
            Return min
        End Using
    End Function

    ''' <summary>
    ''' Returns the maximal element of the given sequence, based on
    ''' the given projection.
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the maximal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current maximal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <returns>The maximal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="selector"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>
    <Extension()> _
    Public Function MaxBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey)) As TSource
        Return source.MaxBy(selector, Comparer(Of TKey).[Default])
    End Function

    ''' <summary>
    ''' Returns the maximal element of the given sequence, based on
    ''' the given projection and the specified comparer for projected values.
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the maximal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current maximal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <param name="comparer">Comparer to use to compare projected values</param>
    ''' <returns>The maximal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/>, <paramref name="selector"/>
    ''' or <paramref name="comparer"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>
    <Extension()> _
    Public Function MaxBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey), comparer As IComparer(Of TKey)) As TSource
        If source Is Nothing Then
            Throw New ArgumentNullException("source")
        End If
        If selector Is Nothing Then
            Throw New ArgumentNullException("selector")
        End If
        If comparer Is Nothing Then
            Throw New ArgumentNullException("comparer")
        End If
        Using sourceIterator = source.GetEnumerator()
            If Not sourceIterator.MoveNext() Then
                Throw New InvalidOperationException("Sequence contains no elements")
            End If
            Dim max = sourceIterator.Current
            Dim maxKey = selector(max)
            While sourceIterator.MoveNext()
                Dim candidate = sourceIterator.Current
                Dim candidateProjected = selector(candidate)
                If comparer.Compare(candidateProjected, maxKey) > 0 Then
                    max = candidate
                    maxKey = candidateProjected
                End If
            End While
            Return max
        End Using
    End Function

    '''<summary>
    ''' Turn the list of objects to a string of Common Seperated Value
    '''</summary>
    '''<param name="source"></param>
    '''<param name="separator"></param>
    '''<typeparam name="T"></typeparam>
    '''<returns></returns>
    ''' <example>
    ''' 	<code>
    ''' 		var values = new[] { 1, 2, 3, 4, 5 };
    '''			string csv = values.ToCSV(';');
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function ToCsv(Of T)(source As IEnumerable(Of T), separator As Char) As String
        If source Is Nothing Then
            Return [String].Empty
        End If
        Dim csv = New StringBuilder()
        source.ForEach(Function(value) csv.AppendFormat("{0}{1}", value, separator))
        Return csv.ToString(0, csv.Length - 1)
    End Function

    '''<summary>
    ''' Turn the list of objects to a string of Common Seperated Value
    '''</summary>
    '''<param name="source"></param>
    '''<typeparam name="T"></typeparam>
    '''<returns></returns>
    ''' <example>
    ''' 	<code>
    ''' 		var values = new[] {1, 2, 3, 4, 5};
    '''			string csv = values.ToCSV();
    ''' 	</code>
    ''' </example>
    <Extension()> _
    Public Function ToCsv(Of T)(source As IEnumerable(Of T)) As String
        Return source.ToCsv(","c)
    End Function

    ''' <summary>
    ''' Returns true if the <paramref name="source"/> is null or without any items.
    ''' </summary>
    <Extension()>
    Public Function IsNullOrEmpty(Of T)(source As IEnumerable(Of T)) As Boolean
        Return (source Is Nothing OrElse Not source.Any())
    End Function

    ''' <summary>
    ''' Returns true if the <paramref name="source"/> is contains at least one item.
    ''' </summary>
    <Extension()>
    Public Function IsNotEmpty(Of T)(source As IEnumerable(Of T)) As Boolean
        Return Not source.IsNullOrEmpty()
    End Function

    ''' <summary>
    ''' Enumerates the sequence and invokes the given action for each value in the sequence.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    <Extension()>
    Public Sub ForEach(Of T)(source As IEnumerable(Of T), onNext As Action(Of T))
        source.AssertNotNull(Function() source)
        onNext.AssertNotNull(Function() onNext)
        For Each item In source
            onNext(item)
        Next
    End Sub

    ''' <summary>
    ''' Enumerates the sequence and invokes the given action for each value in the sequence.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    <Extension()>
    Public Sub ForEach(Of T)(source As IEnumerable(Of T), onNext As Action(Of T, Integer))
        source.AssertNotNull(Function() source)
        onNext.AssertNotNull(Function() onNext)
        Dim i = 0
        For Each item In source
            onNext(item, Math.Max(Threading.Interlocked.Increment(i), i - 1))
        Next
    End Sub

    ''' <summary>
    ''' Returns the on first true.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="source">The source.</param>
    ''' <param name="condition">The condition.</param>
    <Extension()>
    Public Sub ReturnOnFirstTrue(Of T)(source As IEnumerable(Of T), condition As Func(Of T, Boolean))
        source.AssertNotNull(Function() source)
        condition.AssertNotNull(Function() condition)
        For Each item In source
            If condition(item) Then
                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' First matching condition returns T
    ''' </summary>
    <Extension()>
    Public Function FirstMatching(Of T)(source As IEnumerable(Of T), match As Func(Of T, Boolean)) As T
        For Each item In source
            If match(item) Then
                Return item
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Lazily invokes an action for each value in the sequence, and executes an action for successful termination.
    ''' </summary>
    ''' <typeparam name="TSource">Source sequence element type.</typeparam>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    ''' <param name="onCompleted">Action to invoke on successful termination of the sequence.</param>
    ''' <returns>Sequence exhibiting the specified side-effects upon enumeration.</returns>
    <Extension()> _
    Public Function [Do](Of TSource)(source As IEnumerable(Of TSource), onNext As Action(Of TSource), onCompleted As Action) As IEnumerable(Of TSource)
        source.AssertNotNull(Function() source)
        onNext.AssertNotNull(Function() onNext)
        onCompleted.AssertNotNull(Function() onCompleted)
        Return DoHelper(source, onNext, Function(x) x, onCompleted)
    End Function

    ''' <summary>
    ''' Lazily invokes an action for each value in the sequence, and executes an action upon exceptional termination.
    ''' </summary>
    ''' <typeparam name="TSource">Source sequence element type.</typeparam>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    ''' <param name="onError">Action to invoke on exceptional termination of the sequence.</param>
    ''' <returns>Sequence exhibiting the specified side-effects upon enumeration.</returns>
    <Extension()> _
    Public Function [Do](Of TSource)(source As IEnumerable(Of TSource), onNext As Action(Of TSource), onError As Action(Of Exception)) As IEnumerable(Of TSource)
        source.AssertNotNull(Function() source)
        onNext.AssertNotNull(Function() onNext)
        onError.AssertNotNull(Function() onError)
        Return DoHelper(source, onNext, onError, Nothing)
    End Function

    ''' <summary>
    ''' Lazily invokes an action for each value in the sequence, and executes an action upon successful or exceptional termination.
    ''' </summary>
    ''' <typeparam name="TSource">Source sequence element type.</typeparam>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    ''' <param name="onError">Action to invoke on exceptional termination of the sequence.</param>
    ''' <param name="onCompleted">Action to invoke on successful termination of the sequence.</param>
    ''' <returns>Sequence exhibiting the specified side-effects upon enumeration.</returns>
    <Extension()> _
    Public Function [Do](Of TSource)(source As IEnumerable(Of TSource), onNext As Action(Of TSource), onError As Action(Of Exception), onCompleted As Action) As IEnumerable(Of TSource)
        source.AssertNotNull(Function() source)
        onNext.AssertNotNull(Function() onNext)
        onError.AssertNotNull(Function() onError)
        onCompleted.AssertNotNull(Function() onCompleted)
        Return DoHelper(source, onNext, onError, onCompleted)
    End Function

    Private Function DoHelper(Of TSource)(source As IEnumerable(Of TSource), onNext As Action(Of TSource), onError As Action(Of Exception), onCompleted As Action) As IEnumerable(Of TSource)
        Using e = source.GetEnumerator()
            Dim list As IList(Of TSource) = New List(Of TSource)
            While True
                Dim current As TSource = Nothing
                Try
                    If Not e.MoveNext() Then
                        Exit Try
                    End If
                    current = e.Current
                Catch ex As Exception
                    onError(ex)
                    Throw
                End Try
                onNext(current)
                list.Add(current)
            End While
            Return list
            onCompleted()
        End Using
    End Function

End Module

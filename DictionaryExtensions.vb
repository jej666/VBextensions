Imports System.Runtime.CompilerServices
Imports System.Linq
Imports System.Collections.Generic
Imports System.Collections
Imports System.Globalization

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module DictionaryExtensions

    ''' <summary>
    ''' Get the value associated with the key as a int.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetInt(d As IDictionary, key As Object) As Integer
        Return Convert.ToInt32(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Get the value associated with the key as a bool.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetBool(d As IDictionary, key As Object) As Boolean
        Return Convert.ToBoolean(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Get the value associated with the key as a string.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetString(d As IDictionary, key As Object) As String
        Return Convert.ToString(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Get the value associated with the key as a double.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetDouble(d As IDictionary, key As Object) As Double
        Return Convert.ToDouble(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Get the value associated with the key as a datetime.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetDateTime(d As IDictionary, key As Object) As DateTime
        Return Convert.ToDateTime(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Get the value associated with the key as a long.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetLong(d As IDictionary, key As Object) As Long
        Return Convert.ToInt64(d(key), CultureInfo.CurrentCulture)
    End Function

    ''' <summary>
    ''' Returns the value of the first entry found with one of the <paramref name="keys"/> received.
    ''' <para>Returns <paramref name="defaultValue"/> if none of the keys exists in this collection </para>
    ''' </summary>
    ''' <param name="defaultValue">Default value if none of the keys </param>
    ''' <param name="keys"> keys to search for (in order) </param>
    <Extension()> _
    Public Function GetFirstValue(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue), defaultValue As TValue, ParamArray keys As TKey()) As TValue
        For Each key In keys
            If dictionary.ContainsKey(key) Then
                Return dictionary(key)
            End If
        Next
        Return defaultValue
    End Function

    ''' <summary>
    ''' Sorts the specified dictionary.
    ''' </summary>
    ''' <typeparam name="TKey">The type of the key.</typeparam>
    ''' <typeparam name="TValue">The type of the value.</typeparam>
    ''' <param name="dictionary">The dictionary.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function Sort(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue)) As IDictionary(Of TKey, TValue)
        dictionary.AssertNotNull(Function() dictionary)
        Return New SortedDictionary(Of TKey, TValue)(dictionary)
    End Function

    ''' <summary>
    ''' Sorts the specified dictionary.
    ''' </summary>
    ''' <typeparam name="TKey">The type of the key.</typeparam>
    ''' <typeparam name="TValue">The type of the value.</typeparam>
    ''' <param name="dictionary">The dictionary to be sorted.</param>
    ''' <param name="comparer">The comparer used to sort dictionary.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function Sort(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue), comparer As IComparer(Of TKey)) As IDictionary(Of TKey, TValue)
        dictionary.AssertNotNull(Function() dictionary)
        comparer.AssertNotNull(Function() comparer)
        Return New SortedDictionary(Of TKey, TValue)(dictionary, comparer)
    End Function

    ''' <summary>
    ''' Sorts the dictionary by value.
    ''' </summary>
    ''' <typeparam name="TKey">The type of the key.</typeparam>
    ''' <typeparam name="TValue">The type of the value.</typeparam>
    ''' <param name="dictionary">The dictionary.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function SortByValue(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue)) As IDictionary(Of TKey, TValue)
        Return (New SortedDictionary(Of TKey, TValue)(dictionary)).OrderBy(Function(kvp) kvp.Value).ToDictionary(Function(item) item.Key, Function(item) item.Value)
    End Function

    ''' <summary>
    ''' Inverts the specified dictionary. (Creates a new dictionary with the values as key, and key as values)
    ''' </summary>
    ''' <typeparam name="TKey">The type of the key.</typeparam>
    ''' <typeparam name="TValue">The type of the value.</typeparam>
    ''' <param name="dictionary">The dictionary.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function Invert(Of TKey, TValue)(dictionary As IDictionary(Of TKey, TValue)) As IDictionary(Of TValue, TKey)
        dictionary.AssertNotNull(Function() dictionary)
        Return dictionary.ToDictionary(Function(pair) pair.Value, Function(pair) pair.Key)
    End Function

End Module

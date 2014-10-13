Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Globalization

''' <summary>
'''
''' </summary>
<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module StringExtensions

    ''' <summary>
    ''' 	Determines whether the specified string is null or empty.
    ''' </summary>
    ''' <param name = "value">The string value to check.</param>
    <Extension()>
    Public Function IsEmpty(value As String) As Boolean
        Return ((value Is Nothing) OrElse (value.Length = 0))
    End Function

    ''' <summary>
    ''' 	Determines whether the specified string is not null or empty.
    ''' </summary>
    ''' <param name = "value">The string value to check.</param>
    <Extension()>
    Public Function IsNotEmpty(value As String) As Boolean
        Return (value.IsEmpty() = False)
    End Function

    ''' <summary>
    ''' 	Checks whether the string is empty and returns a default value in case.
    ''' </summary>
    ''' <param name = "value">The string to check.</param>
    ''' <param name = "defaultValue">The default value.</param>
    ''' <returns>Either the string or the default value.</returns>
    <Extension()>
    Public Function IfEmpty(value As String, defaultValue As String) As String
        Return (If(value.IsNotEmpty(), value, defaultValue))
    End Function

    ''' <summary>
    ''' 	Formats the value with the quotes using string.Format.
    ''' </summary>
    ''' <param name = "value">The input string.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function WithQuotes(value As String) As String
        Return String.Format(CultureInfo.CurrentCulture, "'{0}'", value)
    End Function

    ''' <summary>
    ''' 	Formats the value with the parameters using string.Format.
    ''' </summary>
    ''' <param name = "value">The input string.</param>
    ''' <param name = "parameters">The parameters.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function FormatWith(value As String, ParamArray parameters As Object()) As String
        Return String.Format(CultureInfo.CurrentCulture, value, parameters)
    End Function

    ''' <summary>
    ''' 	Trims the text to a provided maximum length.
    ''' </summary>
    ''' <param name = "value">The input string.</param>
    ''' <param name = "maxLength">Maximum length.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function TrimToMaxLength(value As String, maxLength As Integer) As String
        Return (If(value Is Nothing OrElse value.Length <= maxLength, value, value.Substring(0, maxLength)))
    End Function

    ''' <summary>
    ''' 	Trims the text to a provided maximum length and adds a suffix if required.
    ''' </summary>
    ''' <param name = "value">The input string.</param>
    ''' <param name = "maxLength">Maximum length.</param>
    ''' <param name = "suffix">The suffix.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function TrimToMaxLength(value As String, maxLength As Integer, suffix As String) As String
        Return (If(value Is Nothing OrElse value.Length <= maxLength, value, String.Concat(value.Substring(0, maxLength), suffix)))
    End Function

    ''' <summary>
    ''' 	Determines whether the comparison value strig is contained within the input value string
    ''' </summary>
    ''' <param name = "inputValue">The input value.</param>
    ''' <param name = "comparisonValue">The comparison value.</param>
    ''' <param name = "comparisonType">Type of the comparison to allow case sensitive or insensitive comparison.</param>
    ''' <returns>
    ''' 	<c>true</c> if input value contains the specified value, otherwise, <c>false</c>.
    ''' </returns>
    <Extension()>
    Public Function Contains(inputValue As String, comparisonValue As String, comparisonType As StringComparison) As Boolean
        Return (inputValue.IndexOf(comparisonValue, comparisonType) <> -1)
    End Function

    ''' <summary>
    ''' 	Reverses / mirrors a string.
    ''' </summary>
    ''' <param name = "value">The string to be reversed.</param>
    ''' <returns>The reversed string</returns>
    <Extension()>
    Public Function Reverse(value As String) As String
        If value.IsEmpty() OrElse (value.Length = 1) Then
            Return value
        End If
        Dim chars = value.ToCharArray()
        Array.Reverse(chars)
        Return New String(chars)
    End Function

    ''' <summary>
    ''' 	Ensures that a string starts with a given prefix.
    ''' </summary>
    ''' <param name = "value">The string value to check.</param>
    ''' <param name = "prefix">The prefix value to check for.</param>
    ''' <returns>The string value including the prefix</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var extension = "txt";
    ''' 		var fileName = string.Concat(file.Name, extension.EnsureStartsWith("."));
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function EnsureStartsWith(value As String, prefix As String) As String
        Return If(value.StartsWith(prefix, StringComparison.CurrentCulture), value, String.Concat(prefix, value))
    End Function

    ''' <summary>
    ''' 	Ensures that a string ends with a given suffix.
    ''' </summary>
    ''' <param name = "value">The string value to check.</param>
    ''' <param name = "suffix">The suffix value to check for.</param>
    ''' <returns>The string value including the suffix</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var url = "http://www.pgk.de";
    ''' 		url = url.EnsureEndsWith("/"));
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function EnsureEndsWith(value As String, suffix As String) As String
        Return If(value.EndsWith(suffix, StringComparison.CurrentCulture), value, String.Concat(value, suffix))
    End Function

    ''' <summary>
    ''' 	Repeats the specified string value as provided by the repeat count.
    ''' </summary>
    ''' <param name = "value">The original string.</param>
    ''' <param name = "repeatCount">The repeat count.</param>
    ''' <returns>The repeated string</returns>
    <Extension()>
    Public Function Repeat(value As String, repeatCount As Integer) As String
        If value.Length = 1 Then
            Return New String(value(0), repeatCount)
        End If
        Dim sb = New StringBuilder(repeatCount * value.Length)
        While Math.Max(Threading.Interlocked.Decrement(repeatCount), repeatCount + 1) > 0
            sb.Append(value)
        End While
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 	Tests whether the contents of a string is a numeric value
    ''' </summary>
    ''' <param name = "value">String to check</param>
    ''' <returns>
    ''' 	Boolean indicating whether or not the string contents are numeric
    ''' </returns>
    ''' <remarks>
    ''' 	Contributed by Kenneth Scott
    ''' </remarks>
    <Extension()>
    Public Function IsNumeric(value As String) As Boolean
        Dim output As Single
        Return Single.TryParse(value, output)
    End Function

    ''' <summary>
    ''' 	Concatenates the specified string value with the passed additional strings.
    ''' </summary>
    ''' <param name = "value">The original value.</param>
    ''' <param name = "values">The additional string values to be concatenated.</param>
    ''' <returns>The concatenated string.</returns>
    <Extension()>
    Public Function ConcatWith(value As String, ParamArray values As String()) As String
        Return String.Concat(value, String.Concat(values))
    End Function

    ''' <summary>
    ''' 	Convert the provided string to a Guid value.
    ''' </summary>
    ''' <param name = "value">The original string value.</param>
    ''' <returns>The Guid</returns>
    <Extension()>
    Public Function ToGuid(value As String) As Guid
        Return New Guid(value)
    End Function

    ''' <summary>
    ''' 	Convert the provided string to a Guid value and returns Guid.Empty if conversion fails.
    ''' </summary>
    ''' <param name = "value">The original string value.</param>
    ''' <returns>The Guid</returns>
    <Extension()>
    Public Function ToGuidSave(value As String) As Guid
        Return value.ToGuidSave(Guid.Empty)
    End Function

    ''' <summary>
    ''' 	Convert the provided string to a Guid value and returns the provided default value if the conversion fails.
    ''' </summary>
    ''' <param name = "value">The original string value.</param>
    ''' <param name = "defaultValue">The default value.</param>
    ''' <returns>The Guid</returns>
    <Extension()>
    Public Function ToGuidSave(value As String, defaultValue As Guid) As Guid
        If value.IsEmpty() Then
            Return defaultValue
        End If
        Try
            Return value.ToGuid()
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>
    ''' 	Gets the string before the given string parameter.
    ''' </summary>
    ''' <param name = "value">The default value.</param>
    ''' <param name = "x">The given string parameter.</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBetween and GetAfter, this does not Trim the result.</remarks>
    <Extension()>
    Public Function GetBefore(value As String, x As String) As String
        Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
        Return If(xPos = -1, [String].Empty, value.Substring(0, xPos))
    End Function

    ''' <summary>
    ''' 	Gets the string between the given string parameters.
    ''' </summary>
    ''' <param name = "value">The source value.</param>
    ''' <param name = "x">The left string sentinel.</param>
    ''' <param name = "y">The right string sentinel</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
    <Extension()>
    Public Function GetBetween(value As String, x As String, y As String) As String
        Dim xPos = value.IndexOf(x, StringComparison.Ordinal)
        Dim yPos = value.LastIndexOf(y, StringComparison.Ordinal)
        If xPos = -1 OrElse xPos = -1 Then
            Return [String].Empty
        End If
        Dim startIndex = xPos + x.Length
        Return If(startIndex >= yPos, [String].Empty, value.Substring(startIndex, yPos - startIndex).Trim())
    End Function

    ''' <summary>
    ''' 	Gets the string after the given string parameter.
    ''' </summary>
    ''' <param name = "value">The default value.</param>
    ''' <param name = "x">The given string parameter.</param>
    ''' <returns></returns>
    ''' <remarks>Unlike GetBefore, this method trims the result</remarks>
    <Extension()>
    Public Function GetAfter(value As String, x As String) As String
        Dim xPos = value.LastIndexOf(x, StringComparison.Ordinal)
        If xPos = -1 Then
            Return [String].Empty
        End If
        Dim startIndex = xPos + x.Length
        Return If(startIndex >= value.Length, [String].Empty, value.Substring(startIndex).Trim())
    End Function

    ''' <summary>
    ''' 	A generic version of System.String.Join()
    ''' </summary>
    ''' <typeparam name = "T">
    ''' 	The type of the array to join
    ''' </typeparam>
    ''' <param name = "separator">
    ''' 	The separator to appear between each element
    ''' </param>
    ''' <param name = "value">
    ''' 	An array of values
    ''' </param>
    ''' <returns>
    ''' 	The join.
    ''' </returns>
    <Extension()>
    Public Function Join(Of T)(separator As String, value As T()) As String
        If value Is Nothing OrElse value.Length = 0 Then
            Return String.Empty
        End If
        If separator Is Nothing Then
            separator = String.Empty
        End If
        Dim converter As Converter(Of T, String) = Function(o) o.ToString()
        Return String.Join(separator, Array.ConvertAll(value, converter))
    End Function

    ''' <summary>
    ''' Remove any instance of the given string pattern from the current string.
    ''' </summary>
    ''' <param name="value">The input.</param>
    ''' <param name="strings">The strings.</param>
    ''' <returns></returns>
    <Extension()>
    Public Function Remove(value As String, ParamArray strings As String()) As String
        Return strings.Aggregate(value, Function(current, c) current.Replace(c, String.Empty))
    End Function

    ''' <summary>Finds out if the specified string contains null, empty or consists only of white-space characters</summary>
    ''' <param name = "value">The input string</param>
    <Extension()>
    Public Function IsEmptyOrWhiteSpace(value As String) As Boolean
        Return (value.IsEmpty() OrElse value.All(Function(t) Char.IsWhiteSpace(t)))
    End Function

    ''' <summary>Determines whether the specified string is not null, empty or consists only of white-space characters</summary>
    ''' <param name = "value">The string value to check</param>
    <Extension()>
    Public Function IsNotEmptyOrWhiteSpace(value As String) As Boolean
        Return (value.IsEmptyOrWhiteSpace() = False)
    End Function

    ''' <summary>Checks whether the string is null, empty or consists only of white-space characters and returns a default value in case</summary>
    ''' <param name = "value">The string to check</param>
    ''' <param name = "defaultValue">The default value</param>
    ''' <returns>Either the string or the default value</returns>
    <Extension()>
    Public Function IfEmptyOrWhiteSpace(value As String, defaultValue As String) As String
        Return (If(value.IsEmptyOrWhiteSpace(), defaultValue, value))
    End Function

    ''' <summary>Uppercase First Letter</summary>
    ''' <param name = "value">The string value to process</param>
    <Extension()>
    Public Function ToUpperFirstLetter(value As String) As String
        If value.IsEmptyOrWhiteSpace() Then
            Return String.Empty
        End If
        Dim valueChars As Char() = value.ToCharArray()
        valueChars(0) = Char.ToUpper(valueChars(0), CultureInfo.CurrentCulture)
        Return New String(valueChars)
    End Function

    ''' <summary>
    ''' Returns the left part of the string.
    ''' </summary>
    ''' <param name="value">The original string.</param>
    ''' <param name="characterCount">The character count to be returned.</param>
    ''' <returns>The left part</returns>
    <Extension()>
    Public Function Left(value As String, characterCount As Integer) As String
        If value Is Nothing Then
            Throw New ArgumentNullException("value")
        End If
        If characterCount >= value.Length Then
            Throw New ArgumentOutOfRangeException("characterCount", characterCount, "characterCount must be less than length of string")
        End If
        Return value.Substring(0, characterCount)
    End Function

    ''' <summary>
    ''' Returns the Right part of the string.
    ''' </summary>
    ''' <param name="value">The original string.</param>
    ''' <param name="characterCount">The character count to be returned.</param>
    ''' <returns>The right part</returns>
    <Extension()>
    Public Function Right(value As String, characterCount As Integer) As String
        If value Is Nothing Then
            Throw New ArgumentNullException("value")
        End If
        If characterCount >= value.Length Then
            Throw New ArgumentOutOfRangeException("characterCount", characterCount, "characterCount must be less than length of string")
        End If
        Return value.Substring(value.Length - characterCount)
    End Function

    ''' <summary>Returns the right part of the string from index.</summary>
    ''' <param name="value">The original value.</param>
    ''' <param name="index">The start index for substringing.</param>
    ''' <returns>The right part.</returns>
    <Extension()>
    Public Function SubstringFrom(value As String, index As Integer) As String
        Return If(index < 0, value, value.Substring(index, value.Length - index))
    End Function

    ''' <summary>
    ''' Returns true if strings are equals, without consideration to case (<see cref="StringComparison.InvariantCultureIgnoreCase"/>)
    ''' </summary>
    <Extension()>
    Public Function CompareWith(s1 As String, s2 As String) As Boolean
        Return String.Compare(s1, s2, True, CultureInfo.CurrentCulture) = 0
    End Function

    ''' <summary>
    ''' Returns true if strings are equals, without consideration to case (<see cref="StringComparison.InvariantCultureIgnoreCase"/>)
    ''' </summary>
    <Extension()>
    Public Function EquivalentTo(s As String, whateverCase As String) As Boolean
        Return String.Equals(s, whateverCase, StringComparison.InvariantCultureIgnoreCase)
    End Function

    ''' <summary>
    ''' Replace all values in string
    ''' </summary>
    ''' <param name="value">The input string.</param>
    ''' <param name="oldValues">List of old values, which must be replaced</param>
    ''' <param name="replacePredicate">Function for replacement old values</param>
    ''' <returns>Returns new string with the replaced values</returns>
    ''' <example>
    ''' 	<code>
    '''         var str = "White Red Blue Green Yellow Black Gray";
    '''         var achromaticColors = new[] {"White", "Black", "Gray"};
    '''         str = str.ReplaceAll(achromaticColors, v => "[" + v + "]");
    '''         // str == "[White] Red Blue Green Yellow [Black] [Gray]"
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function ReplaceAll(value As String, oldValues As IEnumerable(Of String), replacePredicate As Func(Of String, String)) As String
        Dim sbStr = New StringBuilder(value)
        For Each oldValue In oldValues
            Dim newValue = replacePredicate(oldValue)
            sbStr.Replace(oldValue, newValue)
        Next

        Return sbStr.ToString()
    End Function

    ''' <summary>
    ''' Replace all values in string
    ''' </summary>
    ''' <param name="value">The input string.</param>
    ''' <param name="oldValues">List of old values, which must be replaced</param>
    ''' <param name="newValue">New value for all old values</param>
    ''' <returns>Returns new string with the replaced values</returns>
    ''' <example>
    ''' 	<code>
    '''         var str = "White Red Blue Green Yellow Black Gray";
    '''         var achromaticColors = new[] {"White", "Black", "Gray"};
    '''         str = str.ReplaceAll(achromaticColors, "[AchromaticColor]");
    '''         // str == "[AchromaticColor] Red Blue Green Yellow [AchromaticColor] [AchromaticColor]"
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function ReplaceAll(value As String, oldValues As IEnumerable(Of String), newValue As String) As String
        Dim sbStr = New StringBuilder(value)
        For Each oldValue In oldValues
            sbStr.Replace(oldValue, newValue)
        Next

        Return sbStr.ToString()
    End Function

    ''' <summary>
    ''' Replace all values in string
    ''' </summary>
    ''' <param name="value">The input string.</param>
    ''' <param name="oldValues">List of old values, which must be replaced</param>
    ''' <param name="newValues">List of new values</param>
    ''' <returns>Returns new string with the replaced values</returns>
    ''' <example>
    ''' 	<code>
    '''         var str = "White Red Blue Green Yellow Black Gray";
    '''         var achromaticColors = new[] {"White", "Black", "Gray"};
    '''         var exquisiteColors = new[] {"FloralWhite", "Bistre", "DavyGrey"};
    '''         str = str.ReplaceAll(achromaticColors, exquisiteColors);
    '''         // str == "FloralWhite Red Blue Green Yellow Bistre DavyGrey"
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function ReplaceAll(value As String, oldValues As IEnumerable(Of String), newValues As IEnumerable(Of String)) As String
        Dim sbStr = New StringBuilder(value)
        Using newValueEnum = newValues.GetEnumerator()
            For Each old In oldValues
                If Not newValueEnum.MoveNext() Then
                    Throw New ArgumentOutOfRangeException("newValues", "newValues sequence is shorter than oldValues sequence")
                End If
                sbStr.Replace(old, newValueEnum.Current)
            Next
            If newValueEnum.MoveNext() Then
                Throw New ArgumentOutOfRangeException("newValues", "newValues sequence is longer than oldValues sequence")
            End If
        End Using
        Return sbStr.ToString()
    End Function

End Module

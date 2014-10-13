Imports System.Collections.Generic
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Globalization

<HideModuleName(), System.ComponentModel.ImmutableObject(True)> _
Public Module ReflectionExtensions

    ''' <summary>
    ''' Gets the value.
    ''' </summary>
    ''' <param name="member">The member.</param>
    ''' <param name="instance">The instance.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function GetValue(member As MemberInfo, instance As Object) As Object
        Select Case member.MemberType
            Case MemberTypes.[Property]
                Return DirectCast(member, PropertyInfo).GetValue(instance, Nothing)
            Case MemberTypes.Field
                Return DirectCast(member, FieldInfo).GetValue(instance)
            Case Else
                Throw New InvalidOperationException()
        End Select
    End Function

    ''' <summary>
    ''' Sets the value.
    ''' </summary>
    ''' <param name="member">The member.</param>
    ''' <param name="instance">The instance.</param>
    ''' <param name="value">The value.</param>
    ''' <remarks></remarks>
    <Extension()> _
    Public Sub SetValue(member As MemberInfo, instance As Object, value As Object)
        Select Case member.MemberType
            Case MemberTypes.[Property]
                Dim pi = DirectCast(member, PropertyInfo)
                pi.SetValue(instance, value, Nothing)
                Exit Select
            Case MemberTypes.Field
                Dim fi = DirectCast(member, FieldInfo)
                fi.SetValue(instance, value)
                Exit Select
            Case Else
                Throw New InvalidOperationException()
        End Select
    End Sub

    ''' <summary>
    ''' Set object properties on T using the properties collection supplied.
    ''' The properties collection is the collection of "property" to value.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <param name="properties"></param>
    <Extension()> _
    Public Sub SetProperties(Of T As Class)(obj As T, properties As IList(Of KeyValuePair(Of String, String)))
        If obj Is Nothing Then
            Return
        End If
        For Each propVal As KeyValuePair(Of String, String) In properties
            SetProperty(Of T)(obj, propVal.Key, propVal.Value)
        Next
    End Sub

    ''' <summary>
    ''' Set the object properties using the prop name and value
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    <Extension()> _
    Public Sub SetProperty(Of T As Class)(obj As T, propName As String, propVal As String)
        If obj Is Nothing Then
            Throw New ArgumentNullException("Object containing properties to set is null")
        End If
        If propName.IsEmptyOrWhiteSpace Then
            Throw New ArgumentException("Property name not supplied")
        End If
        propName = propName.Trim()
        If String.IsNullOrEmpty(propName) Then
            Throw New ArgumentException("Property name is empty.")
        End If
        Dim type = obj.[GetType]()
        Dim propertyInfo = type.GetProperty(propName)
        If propertyInfo IsNot Nothing AndAlso propertyInfo.CanWrite Then
            If propertyInfo.CanConvertToCorrectType(propVal) Then
                Dim convertedVal As Object = propertyInfo.ConvertToSameType(propVal)
                propertyInfo.SetValue(obj, convertedVal, Nothing)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Set the property value using the string value.
    ''' </summary>
    <Extension()> _
    Public Sub SetProperty(obj As Object, prop As PropertyInfo, propVal As String)
        If prop Is Nothing Then
            Throw New ArgumentNullException("Object containing properties to set is null")
        End If
        If propVal.IsEmptyOrWhiteSpace Then
            Throw New ArgumentException("Property name not supplied")
        End If
        If prop.CanWrite Then
            If prop.CanConvertToCorrectType(propVal) Then
                Dim convertedVal As Object = prop.ConvertToSameType(propVal)
                prop.SetValue(obj, convertedVal, Nothing)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Get all the property values.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="properties"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropertyValues(obj As Object, properties As IList(Of String)) As IList(Of Object)
        Return (From [property] In properties Select propInfo = obj.[GetType]().GetProperty([property]) Select propInfo.GetValue(obj, Nothing)).ToList()
    End Function

    ''' <summary>
    ''' Get all the properties.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetProperties(obj As Object, propsDelimited As String) As IList(Of PropertyInfo)
        Return obj.[GetType]().GetProperties(propsDelimited.Split(","c))
    End Function

    ''' <summary>
    ''' Get all the properties.
    ''' </summary>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetProperties(type As Type, props As String()) As IList(Of PropertyInfo)
        Dim allProps = type.GetProperties()
        Dim propsMap = props.ToDictionary()
        Return (From prop In allProps Where propsMap.ContainsKey(prop.Name)).ToList()
    End Function

    ''' <summary>
    ''' Gets all the properties of the table.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetAllProperties(obj As Object, criteria As Predicate(Of PropertyInfo)) As IList(Of PropertyInfo)
        If obj Is Nothing Then
            Return Nothing
        End If
        Return obj.[GetType]().GetProperties(criteria)
    End Function

    ''' <summary>
    ''' Get the
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetProperties(type As Type, criteria As Predicate(Of PropertyInfo)) As IList(Of PropertyInfo)
        Dim properties = type.GetProperties()
        If properties Is Nothing OrElse properties.Length = 0 Then
            Return Nothing
        End If
        Return (From [property] In properties Let okToAdd = (criteria Is Nothing) OrElse criteria([property]) Where okToAdd Select [property]).ToList()
    End Function

    ''' <summary>
    ''' Gets all the properties of the object as dictionary of property names to propertyInfo.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropertiesAsMap(obj As Object, criteria As Predicate(Of PropertyInfo)) As IDictionary(Of String, PropertyInfo)
        Dim matchedProps = obj.[GetType]().GetProperties(criteria)
        Dim props = New Dictionary(Of String, PropertyInfo)()
        For Each prop As PropertyInfo In matchedProps
            props.Add(prop.Name, prop)
        Next
        Return props
    End Function

    ''' <summary>
    ''' Gets all the properties of the object as dictionary of property names to value of property.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropertiesValues(obj As Object, criteria As Predicate(Of PropertyInfo)) As IDictionary(Of String, Object)
        Dim result = New Dictionary(Of String, Object)()
        Dim props = obj.[GetType]().GetProperties(criteria)
        For Each prop As PropertyInfo In props
            Try
                result.Add(prop.Name, prop.GetValue(obj, Nothing))
            Catch
            End Try
        Next
        Return result
    End Function

    ''' <summary>
    ''' Gets a list of all the writable properties of the class associated with the object.
    ''' </summary>
    ''' <remarks>This method does not take into account, security, generics, etc.
    ''' It only checks whether or not the property can be written to.</remarks>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetWritableProperties(type As Type, criteria As Predicate(Of PropertyInfo)) As IList(Of PropertyInfo)
        Dim props = [type].GetProperties(
            Function([property] As PropertyInfo)
                ' Now determine if it can be added based on criteria.
                Dim okToAdd As Boolean = If((criteria Is Nothing), [property].CanWrite, ([property].CanWrite AndAlso criteria([property])))
                Return okToAdd
            End Function)
        Return props
    End Function

    ''' <summary>
    ''' Invokes the method on the object provided.
    ''' </summary>
    ''' <param name="obj">The object containing the method to invoke</param>
    ''' <param name="methodName">arguments to the method.</param>
    <Extension()> _
    Public Function InvokeMethod(obj As Object, methodName As String, parameters As Object()) As Object
        If methodName.IsEmptyOrWhiteSpace Then
            Throw New ArgumentException("Method name not supplied")
        End If
        methodName = methodName.Trim()
        If String.IsNullOrEmpty(methodName) Then
            Throw New ArgumentException("Method name not provided")
        End If
        Dim method = obj.[GetType]().GetMethod(methodName)
        Dim output As Object = method.Invoke(obj, parameters)
        Return output
    End Function

    ''' <summary>
    ''' Get the property info from the property expression.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="propertyExpression"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function GetProperty(Of T)(obj As Object, propertyExpression As Expression(Of Func(Of T, Object))) As PropertyInfo
        Dim lambda = TryCast(propertyExpression, LambdaExpression)
        Dim memberExpression As MemberExpression
        If TypeOf lambda.Body Is UnaryExpression Then
            Dim unaryExpression = TryCast(lambda.Body, UnaryExpression)
            memberExpression = DirectCast(unaryExpression.Operand, MemberExpression)
        Else
            memberExpression = DirectCast(lambda.Body, MemberExpression)
        End If
        Return TryCast(memberExpression.Member, PropertyInfo)
    End Function

    ''' <summary>
    ''' Get the propertyInfo of the specified property name.
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="propertyName"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetProperty(type As Type, propertyName As String) As PropertyInfo
        Dim props = type.GetProperties(Function([property] As PropertyInfo) String.Compare([property].Name, propertyName, False) = 0)
        Return props(0)
    End Function

    ''' <summary>
    ''' 	Dynamically retrieves a property value.
    ''' </summary>
    ''' <param name = "obj">The object to perform on.</param>
    ''' <param name = "propertyName">The Name of the property.</param>
    ''' <returns>The property value.</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var type = Type.GetType("System.IO.FileInfo, mscorlib");
    ''' 		var file = type.CreateInstance(@"c:\autoexec.bat");
    ''' 		if(file.GetPropertyValue&lt;bool&gt;("Exists")) {
    ''' 		var reader = file.InvokeMethod&lt;StreamReader&gt;("OpenText");
    ''' 		Console.WriteLine(reader.ReadToEnd());
    ''' 		reader.Close();
    ''' 		}
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function GetPropertyValue(obj As Object, propertyName As String) As Object
        Return GetPropertyValue(Of Object)(obj, propertyName, Nothing)
    End Function

    ''' <summary>
    ''' 	Dynamically retrieves a property value.
    ''' </summary>
    ''' <typeparam name = "T">The expected return data type</typeparam>
    ''' <param name = "obj">The object to perform on.</param>
    ''' <param name = "propertyName">The Name of the property.</param>
    ''' <returns>The property value.</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var type = Type.GetType("System.IO.FileInfo, mscorlib");
    ''' 		var file = type.CreateInstance(@"c:\autoexec.bat");
    ''' 		if(file.GetPropertyValue&lt;bool&gt;("Exists")) {
    ''' 		var reader = file.InvokeMethod&lt;StreamReader&gt;("OpenText");
    ''' 		Console.WriteLine(reader.ReadToEnd());
    ''' 		reader.Close();
    ''' 		}
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function GetPropertyValue(Of T)(obj As Object, propertyName As String) As T
        Dim defaultValue As T
        Return GetPropertyValue(obj, propertyName, defaultValue)
    End Function

    ''' <summary>
    ''' 	Dynamically retrieves a property value.
    ''' </summary>
    ''' <typeparam name = "T">The expected return data type</typeparam>
    ''' <param name = "obj">The object to perform on.</param>
    ''' <param name = "propertyName">The Name of the property.</param>
    ''' <param name = "defaultValue">The default value to return.</param>
    ''' <returns>The property value.</returns>
    ''' <example>
    ''' 	<code>
    ''' 		var type = Type.GetType("System.IO.FileInfo, mscorlib");
    ''' 		var file = type.CreateInstance(@"c:\autoexec.bat");
    ''' 		if(file.GetPropertyValue&lt;bool&gt;("Exists")) {
    ''' 		var reader = file.InvokeMethod&lt;StreamReader&gt;("OpenText");
    ''' 		Console.WriteLine(reader.ReadToEnd());
    ''' 		reader.Close();
    ''' 		}
    ''' 	</code>
    ''' </example>
    <Extension()>
    Public Function GetPropertyValue(Of T)(obj As Object, propertyName As String, defaultValue As T) As T
        Dim theType = obj.GetType()
        Dim theProperty = theType.GetProperty(propertyName)
        If theProperty Is Nothing Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Property '{0}' not found.", propertyName), propertyName)
        End If
        Dim value = theProperty.GetValue(obj, Nothing)
        Return If(TypeOf (value) Is T, CType(value, T), defaultValue)
    End Function

    ''' <summary>
    ''' Gets the property value safely, without throwing an exception.
    ''' If an exception is caught, null is returned.
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="propInfo"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropertyValueSafely(obj As Object, propInfo As PropertyInfo) As Object
        If propInfo Is Nothing Then
            Return Nothing
        End If
        Dim result As Object = Nothing
        Try
            result = propInfo.GetValue(obj, Nothing)
        Catch generatedExceptionName As Exception
            result = Nothing
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 	Dynamically sets a property value.
    ''' </summary>
    ''' <param name = "obj">The object to perform on.</param>
    ''' <param name = "propertyName">The Name of the property.</param>
    ''' <param name = "value">The value to be set.</param>
    <Extension()>
    Public Sub SetPropertyValue(obj As Object, propertyName As String, value As Object)
        Dim theType = obj.GetType()
        Dim theProperty = theType.GetProperty(propertyName)
        If theProperty Is Nothing Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Property '{0}' not found.", propertyName), propertyName)
        End If
        If Not theProperty.CanWrite Then
            Throw New ArgumentException(String.Format(CultureInfo.CurrentCulture, "Property '{0}' does not allow writes.", propertyName), propertyName)
        End If
        theProperty.SetValue(obj, value, Nothing)
    End Sub

End Module
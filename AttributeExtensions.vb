Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.CompilerServices

<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module AttributeExtensions

    ''' <summary>
    ''' Gets the attributes of the specified type applied to the class.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj">The obj.</param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetClassAttributes(Of T)(obj As Object) As IList(Of T)
        If obj Is Nothing Then
            Return New List(Of T)()
        End If
        Dim attributes As Object() = obj.[GetType]().GetCustomAttributes(GetType(T), False)
        Return attributes.Cast(Of T)().ToList()
    End Function

    ''' <summary>
    ''' Get a list of property info's that have the supplied attribute applied to it.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropsWithAttributes(Of T As Attribute)(obj As Object) As IDictionary(Of String, KeyValuePair(Of T, PropertyInfo))
        If obj Is Nothing Then
            Return New Dictionary(Of String, KeyValuePair(Of T, PropertyInfo))()
        End If
        Dim map As New Dictionary(Of String, KeyValuePair(Of T, PropertyInfo))()
        Dim props = obj.GetType.GetAllProperties(Nothing)
        For Each prop As PropertyInfo In props
            Dim attrs As Object() = prop.GetCustomAttributes(GetType(T), True)
            If attrs IsNot Nothing AndAlso attrs.Length > 0 Then
                map(prop.Name) = New KeyValuePair(Of T, PropertyInfo)(TryCast(attrs(0), T), prop)
            End If
        Next
        Return map
    End Function

    ''' <summary>
    ''' Get a list of property info's that have the supplied attribute applied to it.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    <Extension()> _
    Public Function GetPropsWithAttributesList(Of T As Attribute)(obj As Object) As IList(Of KeyValuePair(Of T, PropertyInfo))
        If obj Is Nothing Then
            Return New List(Of KeyValuePair(Of T, PropertyInfo))()
        End If
        Dim props = obj.GetType.GetAllProperties(Nothing)
        Return (From prop In props Let attrs = prop.GetCustomAttributes(GetType(T), True)
                Where attrs IsNot Nothing AndAlso attrs.Length > 0
                Select New KeyValuePair(Of T, PropertyInfo)(TryCast(attrs(0), T), prop)).ToList()
    End Function

    ''' <summary>
    ''' 	Gets the first matching attribute defined on the data type.
    ''' </summary>
    ''' <typeparam name = "T">The attribute type tp look for.</typeparam>
    ''' <param name = "obj">The object to look on.</param>
    ''' <returns>The found attribute</returns>
    <Extension()>
    Public Function GetAttribute(Of T As Attribute)(obj As Object) As T
        Return GetAttribute(Of T)(obj, True)
    End Function

    ''' <summary>
    ''' 	Gets the first matching attribute defined on the data type.
    ''' </summary>
    ''' <typeparam name = "T">The attribute type tp look for.</typeparam>
    ''' <param name = "obj">The object to look on.</param>
    ''' <param name = "includeInherited">if set to <c>true</c> includes inherited attributes.</param>
    ''' <returns>The found attribute</returns>
    <Extension()>
    Public Function GetAttribute(Of T As Attribute)(obj As Object, includeInherited As Boolean) As T
        Dim theType = (If(TryCast(obj, Type), obj.GetType()))
        Dim attributes = theType.GetCustomAttributes(theType, includeInherited)
        Return CType(attributes.FirstOrDefault(), T)
    End Function

    ''' <summary>
    ''' 	Gets all matching attribute defined on the data type.
    ''' </summary>
    ''' <typeparam name = "T">The attribute type tp look for.</typeparam>
    ''' <param name = "obj">The object to look on.</param>
    ''' <returns>The found attributes</returns>
    <Extension()>
    Public Function GetAttributes(Of T As Attribute)(obj As Object) As IEnumerable(Of T)
        Return GetAttributes(Of T)(obj, False)
    End Function

    ''' <summary>
    ''' 	Gets all matching attribute defined on the data type.
    ''' </summary>
    ''' <typeparam name = "T">The attribute type tp look for.</typeparam>
    ''' <param name = "obj">The object to look on.</param>
    ''' <param name = "includeInherited">if set to <c>true</c> includes inherited attributes.</param>
    ''' <returns>The found attributes</returns>
    <Extension()>
    Public Function GetAttributes(Of T As Attribute)(obj As Object, includeInherited As Boolean) As IEnumerable(Of T)
        Return (If(TryCast(obj, Type), obj.GetType())).GetCustomAttributes(GetType(T), includeInherited).OfType(Of T)().Select(Function(attribute) attribute)
    End Function

    ''' <summary>
    ''' Retrieves the first defined attribute of the given type from the provider if any.
    ''' </summary>
    ''' <typeparam name="TAttribute">Type of the attribute, which must inherit from <see cref="Attribute"/>.</typeparam>
    ''' <param name="provider" this="true">Any type implementing the interface, which can be an assembly, type, 
    ''' property, method, etc.</param>
    ''' <param name="inherit">Optionally, whether the attribute will be looked in base classes.</param>
    ''' <returns>The attribute instance if defined; <see langword="null"/> otherwise.</returns>
    <Extension()> _
    Public Function GetCustomAttribute(Of TAttribute As Attribute)(provider As ICustomAttributeProvider, Optional inherit As Boolean = True) As TAttribute
        Return GetCustomAttributes(Of TAttribute)(provider, inherit).FirstOrDefault()
    End Function

    ''' <summary>
    ''' Retrieves the first defined attribute of the given type from the provider if any.
    ''' </summary>
    ''' <typeparam name="TAttribute">Type of the attribute, which must inherit from <see cref="Attribute"/>.</typeparam>
    ''' <param name="provider" this="true">Any type implementing the interface, which can be an assembly, type, 
    ''' property, method, etc.</param>
    ''' <param name="inherit">Optionally, whether the attribute will be looked in base classes.</param>
    ''' <returns>The attribute instance if defined; <see langword="null"/> otherwise.</returns>
    <Extension()> _
    Public Function GetCustomAttributes(Of TAttribute As Attribute)(provider As ICustomAttributeProvider, Optional inherit As Boolean = True) As IEnumerable(Of TAttribute)
        provider.AssertNotNull(Function() provider)
        Return provider.GetCustomAttributes(GetType(TAttribute), inherit).Cast(Of TAttribute)()
    End Function

    ''' <summary>
    ''' Determines whether the specified member has attribute.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="member">The member.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function HasAttribute(Of T As Attribute)(member As ICustomAttributeProvider) As Boolean
        Return member.IsDefined(GetType(T), False)
    End Function

    ''' <summary>
    ''' Determines whether the specified member has attribute.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="member">The member.</param>
    ''' <param name="includeInherited">The include inherited.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function HasAttribute(Of T As Attribute)(member As ICustomAttributeProvider, includeInherited As Boolean) As Boolean
        Return member.IsDefined(GetType(T), includeInherited)
    End Function

End Module
Imports System.Configuration
Imports System.Xml

'Builds a custom configuration section in the App.config to specify which tags to query in Process Value Archive database

'Class to extend ConfigurationSection to build section to specify Tags
Public Class TagConfigurationSection
    Inherits ConfigurationSection
    ' 
    'The value of the property here "Tags" needs to match that of the config file section
    '
    '
    'Returns a collection of TagElements
    '
    <ConfigurationProperty("Tags")> _
    Public ReadOnly Property Tags() As TagCollection
        Get
            Return DirectCast(MyBase.Item("Tags"), TagCollection)
        End Get
    End Property
End Class

'
'The collection class that will store the list of each element/item that
'is returned back from the configuration manager.
'
<ConfigurationCollection(GetType(TagElement))> _
Public Class TagCollection
    Inherits ConfigurationElementCollection
    Protected Overrides Function CreateNewElement() As ConfigurationElement
        Return New TagElement()
    End Function

    Protected Overrides Function GetElementKey(element As ConfigurationElement) As Object
        Return DirectCast(element, TagElement).Tag
    End Function

    Default Public Overloads ReadOnly Property Item(idx As Integer) As TagElement
        Get
            Return DirectCast(BaseGet(idx), TagElement)
        End Get
    End Property
End Class

'
'The class that holds onto each element returned by the configuration manager.
'
Public Class TagElement
    Inherits ConfigurationElement
    <ConfigurationProperty("tag", DefaultValue:="", IsKey:=True, IsRequired:=False)> _
    Public Property Tag() As String
        Get
            Return DirectCast(MyBase.Item("tag"), String)
        End Get
        Set(value As String)
            MyBase.Item("tag") = value
        End Set
    End Property
    <ConfigurationProperty("mode", DefaultValue:="", IsKey:=False, IsRequired:=False)> _
    Public Property Mode() As String
        Get
            Return DirectCast(MyBase.Item("mode"), String)
        End Get
        Set(value As String)
            MyBase.Item("mode") = value
        End Set
    End Property
End Class

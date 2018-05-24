Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace RichEditDocumentServerAPIExample.CodeUtils
    Public Class CodeExampleGroup
        Public Sub New()
        End Sub
        Public Property Name() As String
        Public Property Examples() As List(Of CodeExample)
        Public Property Id() As Integer
    End Class

    Public Class CodeExample
        Public Property CodeCS() As String
        Public Property CodeCsHelper() As String
        Public Property CodeVB() As String
        Public Property CodeVbHelper() As String
        Public Property RegionName() As String
        Public Property HumanReadableGroupName() As String
        Public Property ExampleGroup() As String
        Public Property Id() As Integer
    End Class

    Public Class CodeExampleCollection
        Inherits List(Of CodeExample)

        Public Sub Merge(ByVal example As CodeExample)
            Dim item As CodeExample = Me.Find(Function(x) x.HumanReadableGroupName.Equals(example.HumanReadableGroupName) AndAlso x.RegionName.Equals(example.RegionName))
            If item Is Nothing Then
                item = New CodeExample()
                item.HumanReadableGroupName = example.HumanReadableGroupName
                item.RegionName = example.RegionName
                Me.Add(item)
            End If
            item.CodeCS &= example.CodeCS
            item.CodeCsHelper &= example.CodeCsHelper
            item.CodeVB &= example.CodeVB
            item.CodeVbHelper &= example.CodeVbHelper
        End Sub

        Public Sub Merge(ByVal exampleList As List(Of CodeExample))
            For Each item As CodeExample In exampleList
                Me.Merge(item)
            Next item
        End Sub
    End Class


    Public Enum ExampleLanguage
        Csharp = 0
        VB = 1
    End Enum
End Namespace

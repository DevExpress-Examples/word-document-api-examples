Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraTreeList

Namespace RichEditDocumentServerAPIExample.CodeUtils

    Public Class RichEditExample
        Inherits RichEditNode

        Private _Action As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer)

        Public Sub New(ByVal name As String, ByVal codeCS As String, ByVal codeVB As String, ByVal action As Action(Of RichEditDocumentServer), ByVal saveResult As Boolean)
            MyBase.New(name)
            Me.Action = action
            Me.CodeCS = codeCS
            Me.CodeVB = codeVB
            Me.SaveResult = saveResult
        End Sub

        Public Property Action As Action(Of RichEditDocumentServer)
            Get
                Return _Action
            End Get

            Private Set(ByVal value As Action(Of RichEditDocumentServer))
                _Action = value
            End Set
        End Property

        Public Property CodeCS As String

        Public Property CodeVB As String

        Public Property SaveResult As Boolean
    End Class

    Public Class RichEditNode

        Private groupsField As GroupsOfRichEditExamples = New GroupsOfRichEditExamples()

        Public Sub New(ByVal name As String)
            Me.Name = name
        End Sub

        <Browsable(False)>
        Public ReadOnly Property Groups As GroupsOfRichEditExamples
            Get
                Return groupsField
            End Get
        End Property

        Public Property Name As String

        <Browsable(False)>
        Public Property Owner As GroupsOfRichEditExamples
    End Class

    Public Class GroupsOfRichEditExamples
        Inherits BindingList(Of RichEditNode)
        Implements TreeList.IVirtualTreeListData

        Private Sub VirtualTreeGetChildNodes(ByVal info As VirtualTreeGetChildNodesInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetChildNodes
            Dim obj As RichEditNode = TryCast(info.Node, RichEditNode)
            info.Children = obj.Groups
        End Sub

        Protected Overrides Sub InsertItem(ByVal index As Integer, ByVal item As RichEditNode)
            item.Owner = Me
            MyBase.InsertItem(index, item)
        End Sub

        Private Sub VirtualTreeGetCellValue(ByVal info As VirtualTreeGetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetCellValue
            Dim obj As RichEditNode = TryCast(info.Node, RichEditNode)
            Select Case info.Column.Caption
                Case "Name"
                    info.CellData = obj.Name
            End Select
        End Sub

        Private Sub VirtualTreeSetCellValue(ByVal info As VirtualTreeSetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeSetCellValue
        End Sub
    End Class

    Public Class CodeExampleGroup

        Public Property Name As String

        Public Property Examples As List(Of CodeExample)

        Public Property Id As Integer
    End Class

    Public Class CodeExample

        Public Property CodeCS As String

        'public string CodeCsHelper { get; set; }
        Public Property CodeVB As String

        'public string CodeVbHelper { get; set; }
        Public Property RegionName As String

        Public Property HumanReadableGroupName As String

        Public Property ExampleGroup As String

        Public Property Id As Integer
    End Class

    Public Enum ExampleLanguage
        Csharp = 0
        VB = 1
    End Enum
End Namespace

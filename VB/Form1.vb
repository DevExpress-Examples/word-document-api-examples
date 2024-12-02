Imports System.Diagnostics
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraTab
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports RichEditDocumentServerAPIExample.CodeUtils

Namespace RichEditDocumentServerAPIExample

    Public Partial Class Form1
        Inherits DevExpress.XtraEditors.XtraForm

        Private codeEditor As ExampleCodeEditor

        Private wordProcessor As RichEditDocumentServer = New RichEditDocumentServer()

        Private richEditExamples As GroupsOfRichEditExamples = New GroupsOfRichEditExamples()

        Public Sub New()
            InitializeComponent()
            InitExamples()
            ShowExamplesInTreeList(treeList1)
            codeEditor = New ExampleCodeEditor(richEditControlCS, richEditControlVB)
            InitCurrentExampleLanguage()
            InitTreeListControl(richEditExamples)
            ShowFirstExample()
        End Sub

        Private Sub InitExamples()
            Dim examplePath As String = CodeExampleUtils.GetExamplePath("CodeExamples")
            Dim examplesCS As Dictionary(Of String, FileInfo) = CodeExampleUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp)
            Dim examplesVB As Dictionary(Of String, FileInfo) = CodeExampleUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.VB)
            DisableTabs(examplesCS.Count, examplesVB.Count)
            Dim actualExamples As Dictionary(Of String, FileInfo)
            Dim exampleFinder As ExampleFinder
            If examplesCS.Count IsNot 0 Then
                actualExamples = examplesCS
                exampleFinder = New ExampleFinderCSharp()
            Else
                actualExamples = examplesVB
                exampleFinder = New ExampleFinderVB()
            End If

            richEditExamples = CodeExampleUtils.FindExamples(actualExamples, exampleFinder)
        End Sub

        Private Sub InitCurrentExampleLanguage()
            Dim currentLanguage As ExampleLanguage = CodeExampleUtils.DetectExampleLanguage("RichEditDocumentServerAPIExample")
            codeEditor.CurrentExampleLanguage = currentLanguage
            xtraTabControl1.SelectedTabPageIndex = If(currentLanguage = ExampleLanguage.Csharp, 0, 1)
        End Sub

        Private Sub InitTreeListControl(ByVal examples As GroupsOfRichEditExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
        End Sub

        Private Sub ShowExamplesInTreeList(ByVal treeList As TreeList)
#Region "InitializeTreeList"
            treeList.OptionsPrint.UsePrintStyles = True
            treeList.FocusedNodeChanged += New FocusedNodeChangedEventHandler(AddressOf Me.OnNewExampleSelected)
            treeList.OptionsView.ShowColumns = False
            treeList.OptionsView.ShowIndicator = False
#End Region
            Dim col1 As TreeListColumn = New TreeListColumn()
            col1.Caption = "Name"
            col1.VisibleIndex = 0
            col1.OptionsColumn.AllowEdit = False
            col1.OptionsColumn.AllowMove = False
            col1.OptionsColumn.ReadOnly = True
            treeList.Columns.AddRange(New TreeListColumn() {col1})
        End Sub

        Private Sub ShowFirstExample()
            treeList1.ExpandAll()
            If treeList1.Nodes.Count > 0 Then treeList1.FocusedNode = treeList1.MoveFirst().FirstNode
            Dim example As RichEditExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), RichEditExample)
            codeEditor.ShowExample(example)
        End Sub

        Private Sub OnNewExampleSelected(ByVal sender As Object, ByVal e As FocusedNodeChangedEventArgs)
            Dim codeExample As RichEditExample = TryCast(TryCast(sender, TreeList).GetDataRecordByNode(e.Node), RichEditExample)
            If codeExample Is Nothing Then Return
            codeEditor.ShowExample(codeExample)
            codeExampleNameLbl.Text = CodeExampleUtils.ConvertStringToHumanReadableForm(codeExample.Name)
        End Sub

        Private Sub DisableTabs(ByVal examplesCSCount As Integer, ByVal examplesVBCount As Integer)
            If examplesCSCount Is 0 Then
                For Each t As XtraTabPage In xtraTabControl1.TabPages
                    If t.Tag.ToString() Is "CS" Then t.PageEnabled = False
                Next
            End If

            If examplesVBCount Is 0 Then
                For Each t As XtraTabPage In xtraTabControl1.TabPages
                    If t.Tag.ToString() Is "VB" Then t.PageEnabled = False
                Next
            End If
        End Sub

        Private Sub OnRunButtonClick(ByVal sender As Object, ByVal e As EventArgs)
            wordProcessor.CreateNewDocument()
            Dim example As RichEditExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), RichEditExample)
            If example Is Nothing Then Return
            Dim action As Action(Of RichEditDocumentServer) = example.Action
            action(wordProcessor)
            SaveDocumentToFile(example)
        End Sub

        Private Sub SaveDocumentToFile(ByVal example As RichEditExample)
            ' Save the modified document to the file.
            If example.SaveResult Then
                Try
                    wordProcessor.SaveDocument("Result.docx", DocumentFormat.OpenXml)
                    Process.Start(New ProcessStartInfo("Result.docx") With {.UseShellExecute = True})
                Catch __unusedException1__ As Exception
                    MessageBox.Show("Close the Result.docx file.")
                End Try
            End If
        End Sub
    End Class
End Namespace

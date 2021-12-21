Imports System
Imports System.Linq
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module BasicActions

        Public CreateNewDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.CreateNewDocument

        Public LoadDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.LoadDocument

        Public MergeDocumentsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.MergeDocuments

        Public SplitDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.SplitDocument

        Public SaveDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.SaveDocument

        Public PrintDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BasicActions.PrintDocument

        Private Sub CreateNewDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateDocument"
            ' Create a new blank document.
            wordProcessor.CreateNewDocument()
#End Region  ' #CreateDocument
        End Sub

        Private Sub LoadDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#LoadDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
#End Region  ' #LoadDocument
        End Sub

        Private Sub MergeDocuments(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#MergeDocuments"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Insert content from the file at the document end.
            wordProcessor.Document.AppendDocumentContent("Documents//MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
#End Region  ' #MergeDocuments
        End Sub

        Private Sub SplitDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SplitDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Obtain a number of pages in the document.
            Dim pageCount As Integer = wordProcessor.DocumentLayout.GetPageCount()
            ' Check all pages in the document.
            For i As Integer = 0 To pageCount - 1
                ' Access the document page.  
                Dim layoutPage As DevExpress.XtraRichEdit.API.Layout.LayoutPage = wordProcessor.DocumentLayout.GetPage(i)
                ' Access the range of the page's main area.
                Dim mainBodyRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = wordProcessor.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length)
                ' Create the temporary RichEditDocumentServer instance.
                Using tempWordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer = New DevExpress.XtraRichEdit.RichEditDocumentServer()
                    ' Insert the page content to the instance.
                    tempWordProcessor.Document.AppendDocumentContent(mainBodyRange)
                    ' Delete the first empty paragraph.
                    tempWordProcessor.Document.Delete(System.Linq.Enumerable.First(Of DevExpress.XtraRichEdit.API.Native.Paragraph)(tempWordProcessor.Document.Paragraphs).Range)
                    ' Save the document page as an RTF file.
                    Dim fileName As String = System.[String].Format("doc{0}.rtf", i)
                    tempWordProcessor.SaveDocument(fileName, DevExpress.XtraRichEdit.DocumentFormat.Rtf)
                End Using
            Next

            ' Open the File Explorer and select the saved file.
            System.Diagnostics.Process.Start("explorer.exe", "/select," & "doc0.rtf")
#End Region  ' #SplitDocument
        End Sub

        Private Sub SaveDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SaveDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Save the document as a DOCX file.
            wordProcessor.SaveDocument("SavedDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Open the File Explorer and select the saved file.
            System.Diagnostics.Process.Start("explorer.exe", "/select," & "SavedDocument.docx")
#End Region  ' #SaveDocument
        End Sub

        Private Sub PrintDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#PrintDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Print the document to the default printer with the default settings.
            wordProcessor.Print()
#End Region  ' #PrintDocument
        End Sub
    End Module
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports System.Diagnostics
Imports DevExpress.XtraRichEdit.Services
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.Export

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module BasicActions

        Private Sub CreateNewDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateDocument"
            wordProcessor.CreateNewDocument()
#End Region  ' #CreateDocument
        End Sub

        Private Sub LoadDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#LoadDocument"
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
#End Region  ' #LoadDocument
        End Sub

        Private Sub MergeDocuments(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#MergeDocuments"
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            wordProcessor.Document.AppendDocumentContent("Documents//MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
#End Region  ' #MergeDocuments
        End Sub

        Private Sub SplitDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SplitDocument"
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            'Split a document per page
            Dim pageCount As Integer = wordProcessor.DocumentLayout.GetPageCount()
            For i As Integer = 0 To pageCount - 1
                Dim layoutPage As DevExpress.XtraRichEdit.API.Layout.LayoutPage = wordProcessor.DocumentLayout.GetPage(i)
                Dim mainBodyRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = wordProcessor.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length)
                Using tempServer As DevExpress.XtraRichEdit.RichEditDocumentServer = New DevExpress.XtraRichEdit.RichEditDocumentServer()
                    tempServer.Document.AppendDocumentContent(mainBodyRange)
                    'Delete last empty paragraph
                    tempServer.Document.Delete(System.Linq.Enumerable.First(Of DevExpress.XtraRichEdit.API.Native.Paragraph)(tempServer.Document.Paragraphs).Range)
                    'Save the result
                    Dim fileName As String = System.[String].Format("doc{0}.rtf", i)
                    tempServer.SaveDocument(fileName, DevExpress.XtraRichEdit.DocumentFormat.Rtf)
                End Using
            Next

            System.Diagnostics.Process.Start("explorer.exe", "/select," & "doc0.rtf")
#End Region  ' #SplitDocument
        End Sub

        Private Sub SaveDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SaveDocument"
            wordProcessor.Document.AppendDocumentContent("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            wordProcessor.SaveDocument("SavedDocument.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            System.Diagnostics.Process.Start("explorer.exe", "/select," & "SavedDocument.docx")
#End Region  ' #SaveDocument
        End Sub

        Private Sub PrintDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#PrintDocument"
            wordProcessor.Document.AppendDocumentContent("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            wordProcessor.Print()
#End Region  ' #PrintDocument
        End Sub
    End Module
End Namespace

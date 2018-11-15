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
   Public NotInheritable Class BasicActions

       Private Sub New()
       End Sub

        Private Shared Sub CreateNewDocument(ByVal server As RichEditDocumentServer)
'            #Region "#CreateDocument"
            server.CreateNewDocument()
'            #End Region ' #CreateDocument
        End Sub
        Private Shared Sub LoadDocument(ByVal server As RichEditDocumentServer)
'            #Region "#LoadDocument"
            server.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
'            #End Region ' #LoadDocument
        End Sub

        Private Shared Sub MergeDocuments(ByVal server As RichEditDocumentServer)
'       #Region "#MergeDocuments"
        server.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml)
        server.Document.AppendDocumentContent("Documents//MovieRentals.docx", DocumentFormat.OpenXml)
'       #End Region ' #MergeDocuments
        End Sub

        Private Shared Sub SplitDocument(ByVal server As RichEditDocumentServer)
'        #Region "#SplitDocument"
        server.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
        Dim pageCount As Integer = server.DocumentLayout.GetPageCount()

        For i As Integer = 0 To pageCount - 1
            Dim layoutPage As DevExpress.XtraRichEdit.API.Layout.LayoutPage = server.DocumentLayout.GetPage(i)
            Dim mainBodyRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = server.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length)

            Using tempServer As RichEditDocumentServer = New RichEditDocumentServer()
                tempServer.Document.AppendDocumentContent(mainBodyRange)
                tempServer.Document.Delete(tempServer.Document.Paragraphs.First().Range)
                Dim fileName As String = String.Format("doc{0}.rtf", i)
                tempServer.SaveDocument(fileName, DocumentFormat.Rtf)
            End Using
        Next

        System.Diagnostics.Process.Start("explorer.exe", "/select," & "doc0.rtf")
'       #End Region "#SplitDocument"
        End Sub

        Private Shared Sub SaveDocument(ByVal server As RichEditDocumentServer)
'            #Region "#SaveDocument"
            server.Document.AppendDocumentContent("Documents\Grimm.docx", DocumentFormat.OpenXml)
            server.SaveDocument("SavedDocument.docx", DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start("explorer.exe", "/select," & "SavedDocument.docx")
'            #End Region ' #SaveDocument
        End Sub
        Private Shared Sub PrintDocument(ByVal server As RichEditDocumentServer)
'            #Region "#PrintDocument"
            server.Document.AppendDocumentContent("Documents\Grimm.docx", DocumentFormat.OpenXml)
            server.Print()
'            #End Region ' #PrintDocument
        End Sub
   End Class
End Namespace

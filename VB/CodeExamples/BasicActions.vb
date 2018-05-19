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

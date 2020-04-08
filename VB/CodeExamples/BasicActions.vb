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
		Private Sub CreateNewDocument(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CreateDocument"
			wordProcessor.CreateNewDocument()
'			#End Region ' #CreateDocument
		End Sub
		Private Sub LoadDocument(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#LoadDocument"
			wordProcessor.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
'			#End Region ' #LoadDocument
		End Sub
		Private Sub MergeDocuments(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#MergeDocuments"
			wordProcessor.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml)
			wordProcessor.Document.AppendDocumentContent("Documents//MovieRentals.docx",DocumentFormat.OpenXml)
'			#End Region ' #MergeDocuments
		End Sub
		Private Sub SplitDocument(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#SplitDocument"
			wordProcessor.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			'Split a document per page
			Dim pageCount As Integer = wordProcessor.DocumentLayout.GetPageCount()
			For i As Integer = 0 To pageCount - 1
				Dim layoutPage As DevExpress.XtraRichEdit.API.Layout.LayoutPage = wordProcessor.DocumentLayout.GetPage(i)
				Dim mainBodyRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = wordProcessor.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length)
				Using tempServer As New RichEditDocumentServer()
					tempServer.Document.AppendDocumentContent(mainBodyRange)
					'Delete last empty paragraph
					tempServer.Document.Delete(tempServer.Document.Paragraphs.First().Range)
					'Save the result
					Dim fileName As String = String.Format("doc{0}.rtf", i)
					tempServer.SaveDocument(fileName, DocumentFormat.Rtf)
				End Using
			Next i
			System.Diagnostics.Process.Start("explorer.exe", "/select," & "doc0.rtf")
'			#End Region ' #SplitDocument
		End Sub
		Private Sub SaveDocument(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#SaveDocument"
			wordProcessor.Document.AppendDocumentContent("Documents\Grimm.docx", DocumentFormat.OpenXml)
			wordProcessor.SaveDocument("SavedDocument.docx", DocumentFormat.OpenXml)
				System.Diagnostics.Process.Start("explorer.exe", "/select," & "SavedDocument.docx")
'			#End Region ' #SaveDocument
		End Sub
		Private Sub PrintDocument(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#PrintDocument"
			wordProcessor.Document.AppendDocumentContent("Documents\Grimm.docx", DocumentFormat.OpenXml)
			wordProcessor.Print()
'			#End Region ' #PrintDocument
		End Sub
   End Module
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
Imports System.Diagnostics
Imports DevExpress.XtraPrinting
Imports System.IO
Imports DevExpress.XtraRichEdit.Export

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Friend Class ExportActions
		Private Shared Sub SaveImageFromRange(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#SaveImageFromRange"
			Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim docRange As DocumentRange = document.Paragraphs(2).Range
			Dim docImageColl As ReadOnlyDocumentImageCollection = document.Images.Get(docRange)
			If docImageColl.Count > 0 Then
				Dim myImage As DevExpress.Office.Utils.OfficeImage = docImageColl(0).Image
				Dim image As System.Drawing.Image = myImage.NativeImage
				Dim imageName As String = String.Format("Image_at_pos_{0}.png", docRange.Start.ToInt())
				image.Save(imageName)
				System.Diagnostics.Process.Start("explorer.exe", "/select," & imageName)
			End If
'			#End Region ' #SaveImageFromRange
		End Sub

		Private Shared Sub ExportRangeToHtml(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ExportRangeToHtml"
			Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			' Get the range for three paragraphs.
			Dim r As DocumentRange = document.CreateRange(document.Paragraphs(0).Range.Start, document.Paragraphs(0).Range.Length + document.Paragraphs(1).Range.Length + document.Paragraphs(2).Range.Length)
			' Export to HTML.
			Dim htmlText As String = document.GetHtmlText(r, Nothing)
			System.IO.File.WriteAllText("test.html", htmlText)
			' Show the result in a browser window.
			System.Diagnostics.Process.Start("test.html")
'			#End Region ' #ExportRangeToHtml
		End Sub

		Private Shared Sub ExportRangeToPlainText(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ExportRangeToPlainText"
			Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim plainText As String = document.GetText(document.Paragraphs(2).Range)
			System.Windows.Forms.MessageBox.Show(plainText)
'			#End Region ' #ExportRangeToPlainText
		End Sub
		Private Shared Sub ExportToPDF(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ExportToPDF"
			wordProcessor.LoadDocument("Documents\MovieRentals.docx", DocumentFormat.OpenXml)
			'Specify export options:
			Dim options As New PdfExportOptions()
			options.DocumentOptions.Author = "Mark Jones"
			options.Compressed = False
			options.ImageQuality = PdfJpegImageQuality.Highest
			'Export the document to the stream: 
			Using pdfFileStream As New FileStream("Document_PDF.pdf", FileMode.Create)
				wordProcessor.ExportToPdf(pdfFileStream, options)
			End Using
			System.Diagnostics.Process.Start("Document_PDF.pdf")
'			#End Region ' #ExportToPDF
		End Sub
		Private Shared Sub ConvertHTMLtoPDF(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ConvertHTMLtoPDF"
			wordProcessor.LoadDocument("Documents\TextWithImages.htm")
			wordProcessor.ExportToPdf("Document_PDF.pdf")
			System.Diagnostics.Process.Start("Document_PDF.pdf")
'			#End Region ' #ConvertHTMLtoPDF
		End Sub
		Private Shared Sub ConvertHTMLtoDOCX(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ConvertHTMLtoDOCX"
			wordProcessor.LoadDocument("Documents\TextWithImages.htm")
			wordProcessor.SaveDocument("Document_DOCX.docx", DocumentFormat.OpenXml)
			System.Diagnostics.Process.Start("Document_DOCX.docx")
'			#End Region ' #ConvertHTMLtoDOCX
		End Sub
		Private Shared Sub ExportToHTML(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ExportDocumentToHTML"
			wordProcessor.LoadDocument("Documents\MovieRentals.docx", DocumentFormat.OpenXml)
			Dim filePath As String = "Document_HTML.html"
			Using htmlFileStream As New FileStream(filePath, FileMode.Create)
				wordProcessor.SaveDocument(htmlFileStream, DocumentFormat.Html)
			End Using

			System.Diagnostics.Process.Start(filePath)
'			#End Region ' #ExportDocumentToHTML
		End Sub
		Private Shared Sub BeforeExport(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#HandleBeforeExportEvent"
			wordProcessor.LoadDocument("Documents\Grimm.docx")
			AddHandler wordProcessor.BeforeExport, AddressOf BeforeExportHelper.BeforeExport
			wordProcessor.SaveDocument("Document_HTML.html", DocumentFormat.Html)
			System.Diagnostics.Process.Start("Document_HTML.html")
'			#End Region ' #HandleBeforeExportEvent
		End Sub

		#Region "#@HandleBeforeExportEvent"
		Private Class BeforeExportHelper
			Public Shared Sub BeforeExport(ByVal sender As Object, ByVal e As BeforeExportEventArgs)
				Dim options As DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions = TryCast(e.Options, HtmlDocumentExporterOptions)
				If options IsNot Nothing Then
					options.CssPropertiesExportType = DevExpress.XtraRichEdit.Export.Html.CssPropertiesExportType.Link
					options.HtmlNumberingListExportFormat = DevExpress.XtraRichEdit.Export.Html.HtmlNumberingListExportFormat.HtmlFormat
					options.TargetUri = "Document_HTML.html"
				End If
			End Sub
		End Class
		#End Region ' #@HandleBeforeExportEvent
	End Class

End Namespace


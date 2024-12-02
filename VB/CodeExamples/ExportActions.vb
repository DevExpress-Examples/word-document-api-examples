Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports System.Diagnostics
Imports DevExpress.XtraPrinting
Imports System.IO
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Export.Html

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class ExportActions

        Public Shared ExportRangeToHtmlAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ExportRangeToHtml

        Public Shared ExportRangeToPlainTextAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ExportRangeToPlainText

        Public Shared ExportToPDFAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ExportToPDF

        Public Shared ConvertHTMLtoPDFAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ConvertHTMLtoPDF

        Public Shared ConvertHTMLtoDOCXAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ConvertHTMLtoDOCX

        Public Shared ExportToHTMLAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ExportToHTML

        Public Shared BeforeExportAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf BeforeExport

        Private Shared Sub ExportRangeToHtml(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportRangeToHtml"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Paragraphs.Count > 2 Then
                ' Access the range of the first three paragraphs.
                Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(document.Paragraphs(0).Range.Start, document.Paragraphs(2).Range.[End].ToInt() - document.Paragraphs(0).Range.Start.ToInt())
                ' Save text contained in the target range in HTML format.
                Dim htmlText As String = document.GetHtmlText(range, Nothing)
                System.IO.File.WriteAllText("test.html", htmlText)
                ' Show the result in a browser window.
                System.Diagnostics.Process.Start("test.html")
            End If
#End Region  ' #ExportRangeToHtml
        End Sub

        Private Shared Sub ExportRangeToPlainText(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportRangeToPlainText"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Paragraphs.Count > 2 Then
                ' Obtain the plain text contained in the third paragraph. 
                Dim plainText As String = document.GetText(document.Paragraphs(2).Range)
                ' Show the result in a dialog box.
                System.Windows.Forms.MessageBox.Show(plainText)
            End If
#End Region  ' #ExportRangeToPlainText
        End Sub

        Private Shared Sub ExportToPDF(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportToPDF"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Specify PDF export options.
            Dim options As DevExpress.XtraPrinting.PdfExportOptions = New DevExpress.XtraPrinting.PdfExportOptions()
            options.DocumentOptions.Author = "Mark Jones"
            options.Compressed = False
            options.ImageQuality = DevExpress.XtraPrinting.PdfJpegImageQuality.Highest
            ' Export the document to a stream in PDF format. 
            Using pdfFileStream As FileStream = New FileStream("Document_PDF.pdf", FileMode.Create)
                wordProcessor.ExportToPdf(pdfFileStream, options)
            End Using

            ' Show the resulting PDF file. 
            System.Diagnostics.Process.Start("Document_PDF.pdf")
#End Region  ' #ExportToPDF
        End Sub

        Private Shared Sub ConvertHTMLtoPDF(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ConvertHTMLtoPDF"
            ' Load a document from an HTML file.
            wordProcessor.LoadDocument("Documents\TextWithImages.htm")
            ' Save the document as a PDF file.
            wordProcessor.ExportToPdf("Document_PDF.pdf")
            ' Show the resulting PDF file. 
            System.Diagnostics.Process.Start("Document_PDF.pdf")
#End Region  ' #ConvertHTMLtoPDF
        End Sub

        Private Shared Sub ConvertHTMLtoDOCX(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ConvertHTMLtoDOCX"
            ' Load a document from an HTML file.
            wordProcessor.LoadDocument("Documents\TextWithImages.htm")
            ' Save the document as a DOCX file.
            wordProcessor.SaveDocument("Document_DOCX.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Show the resulting DOCX file.
            System.Diagnostics.Process.Start("Document_DOCX.docx")
#End Region  ' #ConvertHTMLtoDOCX
        End Sub

        Private Shared Sub ExportToHTML(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportDocumentToHTML"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Specify the path to the resulting HTML file.
            Dim filePath As String = "Document_HTML.html"
            ' Save the document as an HTML file.
            Using htmlFileStream As FileStream = New FileStream(filePath, FileMode.Create)
                wordProcessor.SaveDocument(htmlFileStream, DevExpress.XtraRichEdit.DocumentFormat.Html)
            End Using

            ' Show the resulting HTML file.
            System.Diagnostics.Process.Start(filePath)
#End Region  ' #ExportDocumentToHTML
        End Sub

        Private Shared Sub BeforeExport(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#HandleBeforeExportEvent"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx")
            ' Handle the Before Export event.
            wordProcessor.BeforeExport += Function(s, e)
                ' Specify the export options before a document is exported to HTML.
                Dim options As DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions = TryCast(e.Options, DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions)
                If options IsNot Nothing Then
                    options.CssPropertiesExportType = DevExpress.XtraRichEdit.Export.Html.CssPropertiesExportType.Link
                    options.HtmlNumberingListExportFormat = DevExpress.XtraRichEdit.Export.Html.HtmlNumberingListExportFormat.HtmlFormat
                    options.TargetUri = "Document_HTML.html"
                End If
            End Function
            ' Save the document as an HTML file.
            wordProcessor.SaveDocument("Document_HTML.html", DevExpress.XtraRichEdit.DocumentFormat.Html)
            ' Show the resulting HTML file.
            System.Diagnostics.Process.Start("Document_HTML.html")
#End Region  ' #HandleBeforeExportEvent
        End Sub
    End Class
End Namespace

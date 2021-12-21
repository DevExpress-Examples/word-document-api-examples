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

        Private Shared Sub SaveImageFromRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SaveImageFromRange"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim docRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((2))).Range
            Dim docImageColl As DevExpress.XtraRichEdit.API.Native.ReadOnlyDocumentImageCollection = document.Images.[Get](docRange)
            If docImageColl.Count > 0 Then
                Dim myImage As DevExpress.Office.Utils.OfficeImage = docImageColl(CInt((0))).Image
                Dim image As System.Drawing.Image = myImage.NativeImage
                Dim imageName As String = System.[String].Format("Image_at_pos_{0}.png", docRange.Start.ToInt())
                image.Save(imageName)
                System.Diagnostics.Process.Start("explorer.exe", "/select," & imageName)
            End If
#End Region  ' #SaveImageFromRange
        End Sub

        Private Shared Sub ExportRangeToHtml(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportRangeToHtml"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Get the range for three paragraphs.
            Dim r As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(document.Paragraphs(CInt((0))).Range.Start, document.Paragraphs(CInt((0))).Range.Length + document.Paragraphs(CInt((1))).Range.Length + document.Paragraphs(CInt((2))).Range.Length)
            ' Export to HTML.
            Dim htmlText As String = document.GetHtmlText(r, Nothing)
            System.IO.File.WriteAllText("test.html", htmlText)
            ' Show the result in a browser window.
            System.Diagnostics.Process.Start("test.html")
#End Region  ' #ExportRangeToHtml
        End Sub

        Private Shared Sub ExportRangeToPlainText(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportRangeToPlainText"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim plainText As String = document.GetText(document.Paragraphs(CInt((2))).Range)
            System.Windows.Forms.MessageBox.Show(plainText)
#End Region  ' #ExportRangeToPlainText
        End Sub

        Private Shared Sub ExportToPDF(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportToPDF"
            wordProcessor.LoadDocument("Documents\MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            'Specify export options:
            Dim options As DevExpress.XtraPrinting.PdfExportOptions = New DevExpress.XtraPrinting.PdfExportOptions()
            options.DocumentOptions.Author = "Mark Jones"
            options.Compressed = False
            options.ImageQuality = DevExpress.XtraPrinting.PdfJpegImageQuality.Highest
            'Export the document to the stream: 
            Using pdfFileStream As System.IO.FileStream = New System.IO.FileStream("Document_PDF.pdf", System.IO.FileMode.Create)
                wordProcessor.ExportToPdf(pdfFileStream, options)
            End Using

            System.Diagnostics.Process.Start("Document_PDF.pdf")
#End Region  ' #ExportToPDF
        End Sub

        Private Shared Sub ConvertHTMLtoPDF(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ConvertHTMLtoPDF"
            wordProcessor.LoadDocument("Documents\TextWithImages.htm")
            wordProcessor.ExportToPdf("Document_PDF.pdf")
            System.Diagnostics.Process.Start("Document_PDF.pdf")
#End Region  ' #ConvertHTMLtoPDF
        End Sub

        Private Shared Sub ConvertHTMLtoDOCX(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ConvertHTMLtoDOCX"
            wordProcessor.LoadDocument("Documents\TextWithImages.htm")
            wordProcessor.SaveDocument("Document_DOCX.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            System.Diagnostics.Process.Start("Document_DOCX.docx")
#End Region  ' #ConvertHTMLtoDOCX
        End Sub

        Private Shared Sub ExportToHTML(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ExportDocumentToHTML"
            wordProcessor.LoadDocument("Documents\MovieRentals.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim filePath As String = "Document_HTML.html"
            Using htmlFileStream As System.IO.FileStream = New System.IO.FileStream(filePath, System.IO.FileMode.Create)
                wordProcessor.SaveDocument(htmlFileStream, DevExpress.XtraRichEdit.DocumentFormat.Html)
            End Using

            System.Diagnostics.Process.Start(filePath)
#End Region  ' #ExportDocumentToHTML
        End Sub

        Private Shared Sub BeforeExport(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#HandleBeforeExportEvent"
            wordProcessor.LoadDocument("Documents\Grimm.docx")
            AddHandler wordProcessor.BeforeExport, AddressOf RichEditDocumentServerAPIExample.CodeExamples.ExportActions.BeforeExportHelper.BeforeExport
            wordProcessor.SaveDocument("Document_HTML.html", DevExpress.XtraRichEdit.DocumentFormat.Html)
            System.Diagnostics.Process.Start("Document_HTML.html")
#End Region  ' #HandleBeforeExportEvent
        End Sub

#Region "#@HandleBeforeExportEvent"
        Private Class BeforeExportHelper

            Public Shared Sub BeforeExport(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.BeforeExportEventArgs)
                Dim options As DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions = TryCast(e.Options, DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions)
                If options IsNot Nothing Then
                    options.CssPropertiesExportType = DevExpress.XtraRichEdit.Export.Html.CssPropertiesExportType.Link
                    options.HtmlNumberingListExportFormat = DevExpress.XtraRichEdit.Export.Html.HtmlNumberingListExportFormat.HtmlFormat
                    options.TargetUri = "Document_HTML.html"
                End If
            End Sub
        End Class
#End Region  ' #@HandleBeforeExportEvent
    End Class
End Namespace

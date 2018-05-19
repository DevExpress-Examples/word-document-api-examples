using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;
using System.Diagnostics;
using DevExpress.XtraPrinting;
using System.IO;
using DevExpress.XtraRichEdit.Export;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ExportActions
    {
        static void SaveImageFromRange(RichEditDocumentServer server)
        {
            #region #SaveImageFromRange
            DevExpress.XtraRichEdit.API.Native.Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            DocumentRange docRange = document.Paragraphs[2].Range;
            ReadOnlyDocumentImageCollection docImageColl = document.Images.Get(docRange);
            if (docImageColl.Count > 0)
            {
                DevExpress.Office.Utils.OfficeImage myImage = docImageColl[0].Image;
                System.Drawing.Image image = myImage.NativeImage;
                string imageName = String.Format("Image_at_pos_{0}.png", docRange.Start.ToInt());
                image.Save(imageName);
                System.Diagnostics.Process.Start("explorer.exe", "/select," + imageName);
            }
            #endregion #SaveImageFromRange
        }

        static void ExportRangeToHtml(RichEditDocumentServer server)
        {
            #region #ExportRangeToHtml
            DevExpress.XtraRichEdit.API.Native.Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            // Get the range for three paragraphs.
            DocumentRange r = document.CreateRange(document.Paragraphs[0].Range.Start, document.Paragraphs[0].Range.Length + document.Paragraphs[1].Range.Length + document.Paragraphs[2].Range.Length);
            // Export to HTML.
            string htmlText = document.GetHtmlText(r, null);
            System.IO.File.WriteAllText("test.html", htmlText);
            // Show the result in a browser window.
            System.Diagnostics.Process.Start("test.html");
            #endregion #ExportRangeToHtml
        }

        static void ExportRangeToPlainText(RichEditDocumentServer server)
        {
            #region #ExportRangeToPlainText
            DevExpress.XtraRichEdit.API.Native.Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            string plainText = document.GetText(document.Paragraphs[2].Range);
            System.Windows.Forms.MessageBox.Show(plainText);
            #endregion #ExportRangeToPlainText
        }
        static void ExportToPDF(RichEditDocumentServer server)
        {
            #region #ExportToPDF
            server.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);
            //Specify export options:
            PdfExportOptions options = new PdfExportOptions();
            options.DocumentOptions.Author = "Mark Jones";
            options.Compressed = false;
            options.ImageQuality = PdfJpegImageQuality.Highest;
            //Export the document to the stream: 
            using (FileStream pdfFileStream = new FileStream("Document_PDF.pdf", FileMode.Create))
            {
                server.ExportToPdf(pdfFileStream, options);
            }
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ExportToPDF
        }
        static void ConvertHTMLtoPDF(RichEditDocumentServer server)
        {
            #region #ConvertHTMLtoPDF
            server.LoadDocument("Documents\\TextWithImages.htm");
            server.ExportToPdf("Document_PDF.pdf");
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ConvertHTMLtoPDF
        }
        static void ConvertHTMLtoDOCX(RichEditDocumentServer server)
        {
            #region #ConvertHTMLtoDOCX
            server.LoadDocument("Documents\\TextWithImages.htm");
            server.SaveDocument("Document_DOCX.docx", DocumentFormat.OpenXml);
            System.Diagnostics.Process.Start("Document_DOCX.docx");
            #endregion #ConvertHTMLtoDOCX
        }
        static void ExportToHTML(RichEditDocumentServer server)
        {
            #region #ExportDocumentToHTML
            server.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);
            string filePath = "Document_HTML.html";
            using (FileStream htmlFileStream = new FileStream(filePath, FileMode.Create))
            {
                server.SaveDocument(htmlFileStream, DocumentFormat.Html);
            }

            System.Diagnostics.Process.Start(filePath);
            #endregion #ExportDocumentToHTML
        }
        static void BeforeExport(RichEditDocumentServer server)
        {
            #region #HandleBeforeExportEvent
            server.LoadDocument("Documents\\Grimm.docx");
            server.BeforeExport += BeforeExportHelper.BeforeExport;
            server.SaveDocument("Document_HTML.html", DocumentFormat.Html);
            System.Diagnostics.Process.Start("Document_HTML.html");
            #endregion #HandleBeforeExportEvent
        }

        #region #@HandleBeforeExportEvent
        class BeforeExportHelper
        {
            public static void BeforeExport(object sender, BeforeExportEventArgs e)
            {
                DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions options = e.Options as HtmlDocumentExporterOptions;
                if (options != null)
                {
                    options.CssPropertiesExportType = DevExpress.XtraRichEdit.Export.Html.CssPropertiesExportType.Link;
                    options.HtmlNumberingListExportFormat = DevExpress.XtraRichEdit.Export.Html.HtmlNumberingListExportFormat.HtmlFormat;
                    options.TargetUri = "Document_HTML.html";
                }
            }
        }
        #endregion #@HandleBeforeExportEvent
    }

}


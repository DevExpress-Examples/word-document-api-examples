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
        static void SaveImageFromRange(RichEditDocumentServer wordProcessor)
        {
            #region #SaveImageFromRange
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;
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

        static void ExportRangeToHtml(RichEditDocumentServer wordProcessor)
        {
            #region #ExportRangeToHtml
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;
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

        static void ExportRangeToPlainText(RichEditDocumentServer wordProcessor)
        {
            #region #ExportRangeToPlainText
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            string plainText = document.GetText(document.Paragraphs[2].Range);
            System.Windows.Forms.MessageBox.Show(plainText);
            #endregion #ExportRangeToPlainText
        }
        static void ExportToPDF(RichEditDocumentServer wordProcessor)
        {
            #region #ExportToPDF
            wordProcessor.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);
            //Specify export options:
            PdfExportOptions options = new PdfExportOptions();
            options.DocumentOptions.Author = "Mark Jones";
            options.Compressed = false;
            options.ImageQuality = PdfJpegImageQuality.Highest;
            //Export the document to the stream: 
            using (FileStream pdfFileStream = new FileStream("Document_PDF.pdf", FileMode.Create))
            {
                wordProcessor.ExportToPdf(pdfFileStream, options);
            }
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ExportToPDF
        }
        static void ConvertHTMLtoPDF(RichEditDocumentServer wordProcessor)
        {
            #region #ConvertHTMLtoPDF
            wordProcessor.LoadDocument("Documents\\TextWithImages.htm");
            wordProcessor.ExportToPdf("Document_PDF.pdf");
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ConvertHTMLtoPDF
        }
        static void ConvertHTMLtoDOCX(RichEditDocumentServer wordProcessor)
        {
            #region #ConvertHTMLtoDOCX
            wordProcessor.LoadDocument("Documents\\TextWithImages.htm");
            wordProcessor.SaveDocument("Document_DOCX.docx", DocumentFormat.OpenXml);
            System.Diagnostics.Process.Start("Document_DOCX.docx");
            #endregion #ConvertHTMLtoDOCX
        }
        static void ExportToHTML(RichEditDocumentServer wordProcessor)
        {
            #region #ExportDocumentToHTML
            wordProcessor.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);
            string filePath = "Document_HTML.html";
            using (FileStream htmlFileStream = new FileStream(filePath, FileMode.Create))
            {
                wordProcessor.SaveDocument(htmlFileStream, DocumentFormat.Html);
            }

            System.Diagnostics.Process.Start(filePath);
            #endregion #ExportDocumentToHTML
        }
        static void BeforeExport(RichEditDocumentServer wordProcessor)
        {
            #region #HandleBeforeExportEvent
            wordProcessor.LoadDocument("Documents\\Grimm.docx");
            wordProcessor.BeforeExport += BeforeExportHelper.BeforeExport;
            wordProcessor.SaveDocument("Document_HTML.html", DocumentFormat.Html);
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


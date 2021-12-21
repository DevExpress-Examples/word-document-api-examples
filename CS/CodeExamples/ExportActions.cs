using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using System.Diagnostics;
using DevExpress.XtraPrinting;
using System.IO;
using DevExpress.XtraRichEdit.Export;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ExportActions
    {
        public static Action<RichEditDocumentServer> SaveImageFromRangeAction = SaveImageFromRange;
        public static Action<RichEditDocumentServer> ExportRangeToHtmlAction = ExportRangeToHtml;
        public static Action<RichEditDocumentServer> ExportRangeToPlainTextAction = ExportRangeToPlainText;
        public static Action<RichEditDocumentServer> ExportToPDFAction = ExportToPDF;
        public static Action<RichEditDocumentServer> ConvertHTMLtoPDFAction = ConvertHTMLtoPDF;
        public static Action<RichEditDocumentServer> ConvertHTMLtoDOCXAction = ConvertHTMLtoDOCX;
        public static Action<RichEditDocumentServer> ExportToHTMLAction = ExportToHTML;
        public static Action<RichEditDocumentServer> BeforeExportAction = BeforeExport;

        static void SaveImageFromRange(RichEditDocumentServer wordProcessor)
        {
            #region #SaveImageFromRange
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;

            // Access the range of the document's third paragraph.
            DocumentRange docRange = document.Paragraphs[2].Range;

            // Obtain all images located in the target range.
            ReadOnlyDocumentImageCollection docImageColl = document.Images.Get(docRange);
            if (docImageColl.Count > 0)
            {
                // Access the first image of the document image collection.
                DevExpress.Office.Utils.OfficeImage myImage = docImageColl[0].Image;

                // Save the image in PNG format. 
                System.Drawing.Image image = myImage.NativeImage;
                string imageName = String.Format("Image_at_pos_{0}.png", docRange.Start.ToInt());
                image.Save(imageName);

                // Open the File Explorer and select the saved image.
                System.Diagnostics.Process.Start("explorer.exe", "/select," + imageName);
            }
            #endregion #SaveImageFromRange
        }

        static void ExportRangeToHtml(RichEditDocumentServer wordProcessor)
        {
            #region #ExportRangeToHtml
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;

            if (document.Paragraphs.Count > 2)
            {
                // Access the range of the first three paragraphs.
                DocumentRange r = document.CreateRange(document.Paragraphs[0].Range.Start, document.Paragraphs[0].Range.Length + document.Paragraphs[1].Range.Length + document.Paragraphs[2].Range.Length);

                // Save text contained in the target range in HTML format.
                string htmlText = document.GetHtmlText(r, null);
                System.IO.File.WriteAllText("test.html", htmlText);

                // Show the result in a browser window.
                System.Diagnostics.Process.Start("test.html");
            }
            #endregion #ExportRangeToHtml
        }

        static void ExportRangeToPlainText(RichEditDocumentServer wordProcessor)
        {
            #region #ExportRangeToPlainText
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            DevExpress.XtraRichEdit.API.Native.Document document = wordProcessor.Document;

            if (document.Paragraphs.Count > 2)
            {
                // Obtain the plain text contained in the third paragraph. 
                string plainText = document.GetText(document.Paragraphs[2].Range);

                // Show the result in a dialog box.
                System.Windows.Forms.MessageBox.Show(plainText);
            }
            #endregion #ExportRangeToPlainText
        }
        static void ExportToPDF(RichEditDocumentServer wordProcessor)
        {
            #region #ExportToPDF
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);

            // Specify PDF export options.
            PdfExportOptions options = new PdfExportOptions();
            options.DocumentOptions.Author = "Mark Jones";
            options.Compressed = false;
            options.ImageQuality = PdfJpegImageQuality.Highest;

            // Export the document to a stream in PDF format. 
            using (FileStream pdfFileStream = new FileStream("Document_PDF.pdf", FileMode.Create))
            {
                wordProcessor.ExportToPdf(pdfFileStream, options);
            }
            
            // Show the resulting PDF file. 
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ExportToPDF
        }
        static void ConvertHTMLtoPDF(RichEditDocumentServer wordProcessor)
        {
            #region #ConvertHTMLtoPDF
            // Load a document from an HTML file.
            wordProcessor.LoadDocument("Documents\\TextWithImages.htm");

            // Save the document as a PDF file.
            wordProcessor.ExportToPdf("Document_PDF.pdf");

            // Show the resulting PDF file. 
            System.Diagnostics.Process.Start("Document_PDF.pdf");
            #endregion #ConvertHTMLtoPDF
        }
        static void ConvertHTMLtoDOCX(RichEditDocumentServer wordProcessor)
        {
            #region #ConvertHTMLtoDOCX
            // Load a document from an HTML file.
            wordProcessor.LoadDocument("Documents\\TextWithImages.htm");

            // Save the document as a DOCX file.
            wordProcessor.SaveDocument("Document_DOCX.docx", DocumentFormat.OpenXml);

            // Show the resulting DOCX file.
            System.Diagnostics.Process.Start("Document_DOCX.docx");
            #endregion #ConvertHTMLtoDOCX
        }
        static void ExportToHTML(RichEditDocumentServer wordProcessor)
        {
            #region #ExportDocumentToHTML
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\MovieRentals.docx", DocumentFormat.OpenXml);

            // Specify the path to the resulting HTML file.
            string filePath = "Document_HTML.html";
            
            // Save the document as an HTML file.
            using (FileStream htmlFileStream = new FileStream(filePath, FileMode.Create))
            {
                wordProcessor.SaveDocument(htmlFileStream, DocumentFormat.Html);
            }
            // Show the resulting HTML file.
            System.Diagnostics.Process.Start(filePath);
            #endregion #ExportDocumentToHTML
        }
        static void BeforeExport(RichEditDocumentServer wordProcessor)
        {
            #region #HandleBeforeExportEvent
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx");

            // Handle the Before Export event.
            wordProcessor.BeforeExport += BeforeExportHelper.BeforeExport;

            // Save the document as an HTML file.
            wordProcessor.SaveDocument("Document_HTML.html", DocumentFormat.Html);
            
            // Show the resulting HTML file.
            System.Diagnostics.Process.Start("Document_HTML.html");
            #endregion #HandleBeforeExportEvent
        }

        class BeforeExportHelper
        {
            public static void BeforeExport(object sender, BeforeExportEventArgs e)
            {
                // Specify the export options before a document is exported to HTML.
                DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions options = e.Options as HtmlDocumentExporterOptions;
                if (options != null)
                {
                    options.CssPropertiesExportType = DevExpress.XtraRichEdit.Export.Html.CssPropertiesExportType.Link;
                    options.HtmlNumberingListExportFormat = DevExpress.XtraRichEdit.Export.Html.HtmlNumberingListExportFormat.HtmlFormat;
                    options.TargetUri = "Document_HTML.html";
                }
            }
        }
    }

}


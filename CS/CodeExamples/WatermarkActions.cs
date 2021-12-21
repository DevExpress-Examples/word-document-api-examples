using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class WatermarkActions
    {
        public static Action<RichEditDocumentServer> CreateTextWatermarkAction = CreateTextWatermark;
        public static Action<RichEditDocumentServer> CreateImageWatermarkAction = CreateImageWatermark;
        static void CreateTextWatermark(RichEditDocumentServer wordProcessor) 
        {
            #region #CreateTextWatermark
            // Access a document.
            Document document = wordProcessor.Document;

            // Check whether the document sections have headers.
            foreach (Section section in document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    // Create an empty header.
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }

            // Specify text watermark options.
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.Color = System.Drawing.Color.LightGray;
            textWatermarkOptions.FontFamily = "Calibri";
            textWatermarkOptions.Layout = WatermarkLayout.Horizontal;
            textWatermarkOptions.Semitransparent = true;

            // Add a text watermark to all document pages.
            document.WatermarkManager.SetText("CONFIDENTIAL", textWatermarkOptions);
            #endregion #CreateTextWatermark
        }
        static void CreateImageWatermark(RichEditDocumentServer wordProcessor) 
        {
            #region #CreateImageWatermark
            //Check whether the document sections have headers.
            foreach (Section section in wordProcessor.Document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    // Create an empty header.
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }
            // Specify image watermark options.
            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Washout = false;
            imageWatermarkOptions.Scale = 2;

            // Add an image watermark to all document pages.
            wordProcessor.Document.WatermarkManager.SetImage(Image.FromFile("Documents//DevExpress.png"), imageWatermarkOptions);
            #endregion #CreateImageWatermark

        }
    }
}

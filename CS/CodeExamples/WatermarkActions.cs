using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RichEditDocumentServerExample.CodeExamples
{
    class WatermarkActions
    {
        static void CreateTextWatermark(RichEditDocumentServer wordProcessor) 
        {
            #region #CreateTextWatermark
            //Check whether the document sections have headers:
            foreach (Section section in wordProcessor.Document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    //If not, create an empty header
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.Color = System.Drawing.Color.LightGray;
            textWatermarkOptions.FontFamily = "Calibri";
            textWatermarkOptions.Layout = WatermarkLayout.Horizontal;
            textWatermarkOptions.Semitransparent = true;

            wordProcessor.Document.WatermarkManager.SetText("CONFIDENTIAL", textWatermarkOptions);
            #endregion #CreateTextWatermark
        }
        static void CreateImageWatermark(RichEditDocumentServer wordProcessor) 
        {
            #region #CreateImageWatermark
            //Check whether the document sections have headers:
            foreach (Section section in wordProcessor.Document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    //If not, create an empty header
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }

            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Washout = false;
            imageWatermarkOptions.Scale = 2;
            wordProcessor.Document.WatermarkManager.SetImage(Image.FromFile("Documents//DevExpress.png"), imageWatermarkOptions);
            #endregion #CreateImageWatermark

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class InlinePicturesActions
    {
        public static Action<RichEditDocumentServer> ImageCollectionAction = ImageCollection;
        public static Action<RichEditDocumentServer> SaveImageToFileAction = SaveImageToFile;

        static void ImageCollection(RichEditDocumentServer wordProcessor)
        {
            #region #ImageCollection
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Obtain all images contained in the document.
            ReadOnlyDocumentImageCollection images = document.Images;

            // If the image width exceeds 50 millimeters, 
            // scale the image proportionally to half its size.
            for (int i = 0; i < images.Count; i++)
            {
                if (images[i].Size.Width > DevExpress.Office.Utils.Units.MillimetersToDocumentsF(50))
                {
                    images[i].ScaleX /= 2;
                    images[i].ScaleY /= 2;
                }
            }
            #endregion #ImageCollection
        }

        static void SaveImageToFile(RichEditDocumentServer wordProcessor)
        {
            #region #SaveImageToFile
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Create a document range.
            DocumentRange myRange = document.CreateRange(0, 100);

            // Obtain all images in the target range.
            ReadOnlyDocumentImageCollection images = document.Images.Get(myRange);
            
            if (images.Count > 0)
            {
                // Save the first retrieved image as a PNG file.
                DevExpress.Office.Utils.OfficeImage myImage = images[0].Image;
                System.Drawing.Image image = myImage.NativeImage;
                string imageName = String.Format("Image_at_pos_{0}.png", images[0].Range.Start.ToInt());
                image.Save(imageName);

                // Open the File Explorer and select the saved image.
                System.Diagnostics.Process.Start("explorer.exe", "/select," + imageName);
            }
            #endregion #SaveImageToFile
        }
    }
}


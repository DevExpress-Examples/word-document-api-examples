using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ShapesActions
    {
       
        static void AddFloatingPicture(RichEditDocumentServer wordProcessor)
        {
            #region #AddFloatingPicture
            Document document = wordProcessor.Document;
            document.AppendText("Line One\nLine Two\nLine Three");
            Shape myPicture = document.Shapes.InsertPicture(document.CreatePosition(15),
                System.Drawing.Image.FromFile("Documents\\beverages.png"));
            myPicture.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            #endregion #AddFloatingPicture
        }

        static void FloatingPictureOffset(RichEditDocumentServer wordProcessor)
        {
            #region #FloatingPictureOffset
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            document.Unit = DevExpress.Office.DocumentUnit.Centimeter;
            Shape myPicture = document.Shapes[1];
            // Clear the qualitative positioning to allow positioning by specifying the numerical offset. 
            myPicture.HorizontalAlignment = ShapeHorizontalAlignment.None;
            myPicture.VerticalAlignment = ShapeVerticalAlignment.None;
            // Specify the reference item for positioning.
            myPicture.RelativeHorizontalPosition = ShapeRelativeHorizontalPosition.LeftMargin;
            myPicture.RelativeVerticalPosition = ShapeRelativeVerticalPosition.TopMargin;
            // Specify the offset value.
            myPicture.Offset = new System.Drawing.PointF(4.5f, 2.0f);
            #endregion #FloatingPictureOffset
        }

        static void ChangeZorderAndWrapping(RichEditDocumentServer wordProcessor)
        {
            #region #ChangeZorderAndWrapping
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            Shape myPicture = document.Shapes[1];
            myPicture.VerticalAlignment = ShapeVerticalAlignment.Top;
            myPicture.ZOrder = document.Shapes[0].ZOrder - 1;
            myPicture.TextWrapping = TextWrappingType.BehindText;
            #endregion #ChangeZorderAndWrapping
        }

        static void AddTextBox(RichEditDocumentServer wordProcessor)
        {
            #region #AddTextBox
            Document document = wordProcessor.Document;
            document.AppendText("Line One\nLine Two\nLine Three");
            Shape myTextBox = document.Shapes.InsertTextBox(document.CreatePosition(15));
            myTextBox.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            // Specify the text box background color.
            myTextBox.Fill.Color = System.Drawing.Color.WhiteSmoke;
            // Draw a border around the text box.
            myTextBox.Line.Color = System.Drawing.Color.Black;
            myTextBox.Line.Thickness = 1;
            // Modify text box content.
            SubDocument textBoxDocument = myTextBox.ShapeFormat.TextBox.Document;
            textBoxDocument.AppendText("TextBox Text");
            CharacterProperties cp = textBoxDocument.BeginUpdateCharacters(textBoxDocument.Range.Start, 7);
            cp.ForeColor = System.Drawing.Color.Orange;
            cp.FontSize = 24;
            textBoxDocument.EndUpdateCharacters(cp);
            #endregion #AddTextBox
        }

        static void InsertRichTextInTextBox(RichEditDocumentServer wordProcessor)
        {
            #region #InsertRichTextInTextBox
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            Shape myTextBox = document.Shapes[0];
            // Allow text box resize to fit contents.
            myTextBox.ShapeFormat.TextBox.HeightRule = TextBoxSizeRule.Auto;
            SubDocument boxedDocument = myTextBox.TextBox.Document;
            int appendPosition = myTextBox.ShapeFormat.TextBox.Document.Range.End.ToInt();
            // Append the second paragraph of the main document to the boxed text.
            DocumentRange newRange = boxedDocument.AppendDocumentContent(document.Paragraphs[1].Range);
            boxedDocument.Paragraphs.Insert(newRange.Start);
            // Insert an image form the main document into the text box.
            boxedDocument.Images.Insert(boxedDocument.CreatePosition(appendPosition), document.Images[0].Image.NativeImage);
            // Resize the image so that its size equals the image in the main document.
            boxedDocument.Images[0].Size = document.Images[0].Size;
            #endregion #InsertRichTextInTextBox
        }

        static void RotateAndResize(RichEditDocumentServer wordProcessor)
        {
            #region #RotateAndResize
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            foreach (Shape s in document.Shapes)
            {
                // Rotate a text box and resize a floating picture.
                if (s.Type == ShapeType.Picture)
                {
                    s.RotationAngle = 45;
                }
                else
                {
                    s.ScaleX = 0.1f;
                    s.ScaleY = 0.1f;
                }
            }
            #endregion #RotateAndResize
        }

        static void SelectShape(RichEditDocumentServer wordProcessor)
        {
            #region #SelectShape
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            document.Selection = document.Shapes[0].Range;
            #endregion #SelectShape
        }
    }
}

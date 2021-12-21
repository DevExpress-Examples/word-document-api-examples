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
        public static Action<RichEditDocumentServer> AddFloatingPictureAction = AddFloatingPicture;
        public static Action<RichEditDocumentServer> FloatingPictureOffsetAction = FloatingPictureOffset;
        public static Action<RichEditDocumentServer> ChangeZorderAndWrappingAction = ChangeZorderAndWrapping;
        public static Action<RichEditDocumentServer> AddTextBoxAction = AddTextBox;
        public static Action<RichEditDocumentServer> InsertRichTextInTextBoxAction = InsertRichTextInTextBox;
        public static Action<RichEditDocumentServer> RotateAndResizeAction = RotateAndResize;
        public static Action<RichEditDocumentServer> SelectShapeAction = SelectShape;

        static void AddFloatingPicture(RichEditDocumentServer wordProcessor)
        {
            #region #AddFloatingPicture
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("Line One\nLine Two\nLine Three");
            
            // Insert a picture at the specified position from the file. 
            Shape myPicture = document.Shapes.InsertPicture(document.CreatePosition(15),
                System.Drawing.Image.FromFile("Documents\\beverages.png"));
            
            // Specify the picture alignment.
            myPicture.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            #endregion #AddFloatingPicture
        }

        static void FloatingPictureOffset(RichEditDocumentServer wordProcessor)
        {
            #region #FloatingPictureOffset
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Specify the document's measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Centimeter;

            if (document.Shapes.Count > 1)
            {
                // Access a picture.
                Shape myPicture = document.Shapes[1];

                // Clear the horizontal and vertical alignment values.
                myPicture.HorizontalAlignment = ShapeHorizontalAlignment.None;
                myPicture.VerticalAlignment = ShapeVerticalAlignment.None;

                // The picture's horizontal position is relative to the left margin.
                myPicture.RelativeHorizontalPosition = ShapeRelativeHorizontalPosition.LeftMargin;
                // The picture's vertical position is relative to the top margin.
                myPicture.RelativeVerticalPosition = ShapeRelativeVerticalPosition.TopMargin;

                // Specify the offset value.
                myPicture.Offset = new System.Drawing.PointF(4.5f, 2.0f);
            }
            #endregion #FloatingPictureOffset
        }

        static void ChangeZorderAndWrapping(RichEditDocumentServer wordProcessor)
        {
            #region #ChangeZorderAndWrapping
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Shapes.Count > 1)
            {
                // Access a picture.
                Shape myPicture = document.Shapes[1];

                // Align the picture vertically.
                myPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

                // Specify the picture position in the z-order.
                myPicture.ZOrder = document.Shapes[0].ZOrder - 1;

                // Display document text over the picture.
                myPicture.TextWrapping = TextWrappingType.BehindText;
            }
            #endregion #ChangeZorderAndWrapping
        }

        static void AddTextBox(RichEditDocumentServer wordProcessor)
        {
            #region #AddTextBox
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("Line One\nLine Two\nLine Three");

            // Insert a text box at the specified position.
            Shape myTextBox = document.Shapes.InsertTextBox(document.CreatePosition(15));
            
            // Align the text box horizontally.
            myTextBox.HorizontalAlignment = ShapeHorizontalAlignment.Center;

            // Specify the text box background color.
            myTextBox.Fill.Color = System.Drawing.Color.WhiteSmoke;

            // Draw a border around the text box.
            myTextBox.Line.Color = System.Drawing.Color.Black;
            myTextBox.Line.Thickness = 1;
            
            // Modify text box content.
            SubDocument textBoxDocument = myTextBox.ShapeFormat.TextBox.Document;
            textBoxDocument.AppendText("TextBox Text");

            // Format the boxed text.
            CharacterProperties cp = textBoxDocument.BeginUpdateCharacters(textBoxDocument.Range.Start, 7);
            cp.ForeColor = System.Drawing.Color.Orange;
            cp.FontSize = 24;
            textBoxDocument.EndUpdateCharacters(cp);
            #endregion #AddTextBox
        }

        static void InsertRichTextInTextBox(RichEditDocumentServer wordProcessor)
        {
            #region #InsertRichTextInTextBox
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access a text box.
            Shape myTextBox = document.Shapes[0];

            // Allow text box resize to fit contents.
            myTextBox.ShapeFormat.TextBox.HeightRule = TextBoxSizeRule.Auto;
            SubDocument boxedDocument = myTextBox.ShapeFormat.TextBox.Document;
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
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Check all shapes in the document.
            foreach (Shape s in document.Shapes)
            {
                // Rotate pictures.
                if (s.Type == ShapeType.Picture)
                {
                    s.RotationAngle = 45;
                }
                // Resize text boxes.
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
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Shapes.Count > 1)
            {
                // Select the second drawing object in the shape collection.
                document.Selection = document.Shapes[1].Range;
            }
            #endregion #SelectShape
        }
    }
}

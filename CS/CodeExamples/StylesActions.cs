using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class StylesAction
    {
        public static Action<RichEditDocumentServer> CreateNewCharacterStyleAction = CreateNewCharacterStyle;
        public static Action<RichEditDocumentServer> CreateNewParagraphStyleAction = CreateNewParagraphStyle;
        public static Action<RichEditDocumentServer> CreateNewLinkedStyleAction = CreateNewLinkedStyle;

        static void CreateNewCharacterStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewCharacterStyle
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the character style with the specified name.
            CharacterStyle cstyle = document.CharacterStyles["MyCStyle"];

            // If the style with the specified name does not exist
            // create a new character style and specify the style settings.
            if (cstyle == null)
            {
                cstyle = document.CharacterStyles.CreateNew();
                cstyle.Name = "MyCStyle";
                cstyle.Parent = document.CharacterStyles["Default Paragraph Font"];
                cstyle.ForeColor = System.Drawing.Color.DarkOrange;
                cstyle.Strikeout = StrikeoutType.Double;
                cstyle.FontName = "Verdana";
                // Add the style to the collection of character styles.
                document.CharacterStyles.Add(cstyle);
            }

            // Access the range of the first paragraph.
            DocumentRange myRange = document.Paragraphs[0].Range;
            
            // Access character formatting of the target range.
            CharacterProperties charProps =
                document.BeginUpdateCharacters(myRange);
            
            // Apply the created character style to the target range.
            charProps.Style = cstyle;

            // Finalize to modify character formatting.
            document.EndUpdateCharacters(charProps);
            #endregion #CreateNewCharacterStyle
        }

        static void CreateNewParagraphStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewParagraphStyle
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access a paragraph style with the specified name.
            ParagraphStyle pstyle = document.ParagraphStyles["MyPStyle"];

            // If the style with the specified name does not exist
            // create a new paragraph style and specify the style settings.
            if (pstyle == null)
            {
                pstyle = document.ParagraphStyles.CreateNew();
                pstyle.Name = "MyPStyle";
                pstyle.LineSpacingType = ParagraphLineSpacing.Double;
                pstyle.Alignment = ParagraphAlignment.Center;
                // Add the style to the collection of paragraph styles.
                document.ParagraphStyles.Add(pstyle);
            }

            if (document.Paragraphs.Count > 2)
            {
                // Apply the created paragraph style to the third document paragraph.
                document.Paragraphs[2].Style = pstyle;
            }
            #endregion #CreateNewParagraphStyle
        }

        static void CreateNewLinkedStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewLinkedStyle
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText("Line One\nLine Two\nLine Three");

            // Finalize to edit the document.
            document.EndUpdate();

            // Access a paragraph style with the specified name.
            ParagraphStyle lstyle = document.ParagraphStyles["MyLinkedStyle"];

            // If the style with the specified name does not exist
            // create a new paragraph and character styles and specify their settings.
            if (lstyle == null)
            {
                // Start to edit the document.
                document.BeginUpdate();

                // Create a paragraph style and specify its settings.
                lstyle = document.ParagraphStyles.CreateNew();
                lstyle.Name = "MyLinkedStyle";
                lstyle.LineSpacingType = ParagraphLineSpacing.Double;
                lstyle.Alignment = ParagraphAlignment.Center;
                document.ParagraphStyles.Add(lstyle);

                // Create a character style and specify its settings.
                CharacterStyle lcstyle = document.CharacterStyles.CreateNew();
                lcstyle.Name = "MyLinkedCStyle";
                document.CharacterStyles.Add(lcstyle);

                // Set the created character style to the created paragraph style.
                lcstyle.LinkedStyle = lstyle;

                // Specify the created character style's settings.
                lcstyle.ForeColor = System.Drawing.Color.DarkGreen;
                lcstyle.Strikeout = StrikeoutType.Single;
                lcstyle.FontSize = 24;

                // Finalize to edit the document.
                document.EndUpdate();

                // Save the resulting document and select it in the File Explorer.
                document.SaveDocument("LinkedStyleSample.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start("explorer.exe", "/select," + "LinkedStyleSample.docx");
            }
            #endregion #CreateNewLinkedStyle
        }
    }
}

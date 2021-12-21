using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Drawing;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class FormattingActions
    {
        public static Action<RichEditDocumentServer> FormatTextAction = FormatText;
        public static Action<RichEditDocumentServer> ChangeSpacingAction = ChangeSpacing;
        public static Action<RichEditDocumentServer> ResetCharacterFormattingAction = ResetCharacterFormatting;
        public static Action<RichEditDocumentServer> FormatParagraphAction = FormatParagraph;
        public static Action<RichEditDocumentServer> ResetParagraphFormattingAction = ResetParagraphFormatting;


        static void FormatText(RichEditDocumentServer wordProcessor)
        {
            #region #FormatText
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText("Normal\nFormatted\nNormal");
            
            // Finalize to edit the document.
            document.EndUpdate();
            
            // Access the range of the document's second paragraph.
            DocumentRange range = document.Paragraphs[1].Range;

            // Start to modify character formatting of the target range.
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            
            // Specify character formatting options.
            cp.FontName = "Comic Sans MS";
            cp.FontSize = 18;
            cp.ForeColor = Color.Blue;
            cp.BackColor = Color.Snow;
            cp.Underline = UnderlineType.DoubleWave;
            cp.UnderlineColor = Color.Red;
            
            // Finalize to modify character formatting.
            document.EndUpdateCharacters(cp);
            #endregion #FormatText
        }

        static void ChangeSpacing(RichEditDocumentServer wordProcessor) 
        {
            #region #ChangeCharacterSpacing
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText("Normal\nFormatted\nNormal");
            
            // Finalize to edit the document.
            document.EndUpdate();

            // Access the range of the document's second paragraph.
            DocumentRange range = document.Paragraphs[1].Range;

            // Start to modify character formatting of the target range.
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            
            // Change character spacing and scaling.
            cp.Scale = 150;
            cp.Spacing = -2;
            // Raise the text by 2 points.
            cp.Position = 2;

            // Finalize to modify character formatting.
            document.EndUpdateCharacters(cp);
            #endregion #ChangeCharacterSpacing
        }

        static void ResetCharacterFormatting(RichEditDocumentServer wordProcessor)
        {
            #region #ResetCharacterFormatting
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;            

            // Access the range of the document's first paragraph.
            DocumentRange range = document.Paragraphs[0].Range;

            // Start to modify character formatting of the target range.
            CharacterProperties cp = document.BeginUpdateCharacters(range);

            // Set the font size and font name of the target range's characters to default values.   
            // Other character properties remain intact.
            cp.Reset(CharacterPropertiesMask.FontSize | CharacterPropertiesMask.FontName | CharacterPropertiesMask.FontNameAscii);

            // Finalize to modify character formatting.
            document.EndUpdateCharacters(cp);
            #endregion #ResetCharacterFormatting
        }
        static void FormatParagraph(RichEditDocumentServer wordProcessor)
        {
            #region #FormatParagraph
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText("Modified Paragraph\nNormal\nNormal");
            
            // Finalize to edit the document.
            document.EndUpdate();

            // Access the first paragraph range.
            DocumentRange range = document.Paragraphs[0].Range;

            // Start to edit the paragraph.
            ParagraphProperties pp = document.BeginUpdateParagraphs(range);

            // Specify the paragraph's alignment.
            pp.Alignment = ParagraphAlignment.Center;

            // Specify the paragraph's line spacing.
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 3;

            // Set the paragraph’s left indent to 0.5 document unit.
            // Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5f);
            
            // Start to modify tab stops in the paragraph.
            TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);

            // Create a new tab stop for the paragraph.
            TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();

            // Specify the tab stop's alignment type.
            tbi.Alignment = TabAlignmentType.Center;

            // Set the tab stop position to 1.5 document unit.
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            
            // Add the tab stop to the collection of tab stops.
            tbiColl.Add(tbi);
            
            // Finalize to modify tab stops in the paragraph.
            pp.EndUpdateTabs(tbiColl);

            // Finalize to edit the paragraph.
            document.EndUpdateParagraphs(pp);
            #endregion #FormatParagraph
        }
        static void ResetParagraphFormatting(RichEditDocumentServer wordProcessor)
        {
            #region #ResetParagraphFormatting
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the range of the document's first paragraph.
            DocumentRange range = document.Paragraphs[0].Range;
            
            // Start to edit the paragraph.
            ParagraphProperties cp = document.BeginUpdateParagraphs(range);

            // Set alignmment and first line indent of the target paragraph to default values.   
            // Other paragraph properties remain intact.
            cp.Reset(ParagraphPropertiesMask.Alignment | ParagraphPropertiesMask.FirstLineIndent);
            
            // Finalize to edit the paragraph.
            document.EndUpdateParagraphs(cp);
            #endregion #ResetParagraphFormatting
        }
    }
}

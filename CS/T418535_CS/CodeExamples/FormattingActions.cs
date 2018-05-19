using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class FormattingActions
    {
        
        static void FormatText(RichEditDocumentServer server)
        {
            #region #FormatText
            Document document = server.Document;
            document.BeginUpdate();
            document.AppendText("Normal\nFormatted\nNormal");
            document.EndUpdate();
            DocumentRange range = document.Paragraphs[1].Range;
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.FontName = "Comic Sans MS";
            cp.FontSize = 18;
            cp.ForeColor = Color.Blue;
            cp.BackColor = Color.Snow;
            cp.Underline = UnderlineType.DoubleWave;
            cp.UnderlineColor = Color.Red;
            document.EndUpdateCharacters(cp);
            #endregion #FormatText
        }
        static void ResetCharacterFormatting(RichEditDocumentServer server)
        {
            #region #ResetCharacterFormatting
            Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            // Set font size and font name of the characters in the first paragraph to default. 
            // Other character properties remain intact.
            DocumentRange range = document.Paragraphs[0].Range;
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.Reset(CharacterPropertiesMask.FontSize | CharacterPropertiesMask.FontName);
            document.EndUpdateCharacters(cp);
            #endregion #ResetCharacterFormatting
        }
        static void FormatParagraph(RichEditDocumentServer server)
        {
            #region #FormatParagraph
            Document document = server.Document;
            document.BeginUpdate();
            document.AppendText("Modified Paragraph\nNormal\nNormal");
            document.EndUpdate();
            DocumentPosition pos = document.Range.Start;
            DocumentRange range = document.CreateRange(pos, 0);
            ParagraphProperties pp = document.BeginUpdateParagraphs(range);
            // Center paragraph
            pp.Alignment = ParagraphAlignment.Center;
            // Set triple spacing
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 3;
            // Set left indent at 0.5".
            // Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5f);
            // Set tab stop at 1.5"
            TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
            TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            tbi.Alignment = TabAlignmentType.Center;
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            tbiColl.Add(tbi);
            pp.EndUpdateTabs(tbiColl);
            document.EndUpdateParagraphs(pp);
            #endregion #FormatParagraph
        }
        static void ResetParagraphFormatting(RichEditDocumentServer server)
        {
            #region #ResetParagraphFormatting
            Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            // Set alignment and indentation of the first line in the first paragraph to default. 
            // Other paragraph properties remain intact.
            DocumentRange range = document.Paragraphs[0].Range;
            ParagraphProperties cp = document.BeginUpdateParagraphs(range);
            cp.Reset(ParagraphPropertiesMask.Alignment | ParagraphPropertiesMask.FirstLineIndent);
            document.EndUpdateParagraphs(cp);
            #endregion #ResetParagraphFormatting
        }
    }
}

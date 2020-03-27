using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RTEDocumentServerExamples.CodeExamples
{
    class StylesAction
    {   
        static void CreateNewCharacterStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewCharacterStyle
            Document document = wordProcessor.Document;
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            CharacterStyle cstyle = document.CharacterStyles["MyCStyle"];
            if (cstyle == null)
            {
                cstyle = document.CharacterStyles.CreateNew();
                cstyle.Name = "MyCStyle";
                cstyle.Parent = document.CharacterStyles["Default Paragraph Font"];
                cstyle.ForeColor = System.Drawing.Color.DarkOrange;
                cstyle.Strikeout = StrikeoutType.Double;
                cstyle.FontName = "Verdana";
                document.CharacterStyles.Add(cstyle);
            }
            DocumentRange myRange = document.Paragraphs[0].Range;
            CharacterProperties charProps =
                document.BeginUpdateCharacters(myRange);
            charProps.Style = cstyle;
            document.EndUpdateCharacters(charProps);
            #endregion #CreateNewCharacterStyle
        }

        static void CreateNewParagraphStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewParagraphStyle
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            ParagraphStyle pstyle = document.ParagraphStyles["MyPStyle"];
            if (pstyle == null)
            {
                pstyle = document.ParagraphStyles.CreateNew();
                pstyle.Name = "MyPStyle";
                pstyle.LineSpacingType = ParagraphLineSpacing.Double;
                pstyle.Alignment = ParagraphAlignment.Center;
                document.ParagraphStyles.Add(pstyle);
            }
            document.Paragraphs[2].Style = pstyle;
            #endregion #CreateNewParagraphStyle
        }

        static void CreateNewLinkedStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNewLinkedStyle
            Document document = wordProcessor.Document;
            document.BeginUpdate();
            document.AppendText("Line One\nLine Two\nLine Three");
            document.EndUpdate();
            ParagraphStyle lstyle = document.ParagraphStyles["MyLinkedStyle"];
            if (lstyle == null)
            {
                document.BeginUpdate();
                lstyle = document.ParagraphStyles.CreateNew();
                lstyle.Name = "MyLinkedStyle";
                lstyle.LineSpacingType = ParagraphLineSpacing.Double;
                lstyle.Alignment = ParagraphAlignment.Center;
                document.ParagraphStyles.Add(lstyle);

                CharacterStyle lcstyle = document.CharacterStyles.CreateNew();
                lcstyle.Name = "MyLinkedCStyle";
                document.CharacterStyles.Add(lcstyle);
                lcstyle.LinkedStyle = lstyle;

                lcstyle.ForeColor = System.Drawing.Color.DarkGreen;
                lcstyle.Strikeout = StrikeoutType.Single;
                lcstyle.FontSize = 24;
                document.EndUpdate();
                document.SaveDocument("LinkedStyleSample.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start("explorer.exe", "/select," + "LinkedStyleSample.docx");
            }
            #endregion #CreateNewLinkedStyle
        }
    }
}

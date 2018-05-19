using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class RangeActions
    {       
        static void SelectTextInRange(RichEditDocumentServer server)
        {
            #region #SelectTextInRange
            Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            DocumentPosition myStart = document.CreatePosition(69);
            DocumentRange myRange = document.CreateRange(myStart, 216);
            document.Selection = myRange;
            #endregion #SelectTextInRange
        }

        static void InsertTextAtCaretPosition(RichEditDocumentServer server)
        {
            #region #InsertTextAtCaretPosition
            Document document = server.Document;
            DocumentPosition pos = document.CaretPosition;
            SubDocument doc = pos.BeginUpdateDocument();
            doc.InsertText(pos, " INSERTED TEXT ");
            pos.EndUpdateDocument(doc);
            #endregion #InsertTextAtCaretPosition
        }

        static void InsertTextInRange(RichEditDocumentServer server)
        {
            #region #InsertTextInRange
            Document document = server.Document;
            document.AppendText("ABCDEFGH");
            DocumentRange r1 = document.CreateRange(1, 3);
            DocumentPosition pos1 = document.CreatePosition(2);
            DocumentRange r2 = document.InsertText(pos1, ">>NewText<<");
            string s1 = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End);
            string s2 = String.Format("Range r2 starts at {0}, ends at {1}", r2.Start, r2.End);
            document.Paragraphs.Append();
            document.AppendText(s1);
            document.Paragraphs.Append();
            document.AppendText(s2);
            #endregion #InsertTextInRange
        }

        static void AppendTextToRange(RichEditDocumentServer server)
        {
            #region #AppendTextToRange
            Document document = server.Document;
            document.AppendText("abcdefgh");
            DocumentRange r1 = document.AppendText("X");
            string s1 = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End);
            document.AppendText("Y");
            document.AppendText("Z");
            string s2 = String.Format("Currently range r1 starts at {0}, ends at {1}", r1.Start, r1.End);
            document.Paragraphs.Append();
            document.AppendText(s1);
            document.Paragraphs.Append();
            document.AppendText(s2);
            #endregion #AppendTextToRange
        }

        static void CopyAndPasteRange(RichEditDocumentServer server)
        {
            #region #CopyAndPasteRange
            Document document = server.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            DocumentRange myRange = document.Paragraphs[0].Range;
            document.Copy(myRange);
            document.Paste(DocumentFormat.PlainText);
            #endregion #CopyAndPasteRange
        }

        static void AppendToParagraph(RichEditDocumentServer server)
        {
            #region #AppendToParagraph
            Document document = server.Document;
            document.BeginUpdate();
            document.AppendText("First Paragraph\nSecond Paragraph\nThird Paragraph");
            document.EndUpdate();
            DocumentPosition pos = document.CaretPosition;
            SubDocument doc = pos.BeginUpdateDocument();
            Paragraph par = doc.Paragraphs.Get(pos);
            DocumentPosition newPos = doc.CreatePosition(par.Range.End.ToInt() - 1);
            doc.InsertText(newPos, "<<Appended to Paragraph End>>");
            pos.EndUpdateDocument(doc);
            #endregion #AppendToParagraph
        }
    }
}


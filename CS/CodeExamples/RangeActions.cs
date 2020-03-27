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
        static void SelectTextInRange(RichEditDocumentServer wordProcessor)
        {
            #region #SelectTextInRange
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            DocumentPosition myStart = document.CreatePosition(69);
            DocumentRange myRange = document.CreateRange(myStart, 216);
            document.Selection = myRange;
            #endregion #SelectTextInRange
        }        

        static void InsertTextInRange(RichEditDocumentServer wordProcessor)
        {
            #region #InsertTextInRange
            Document document = wordProcessor.Document;
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

        static void AppendTextToRange(RichEditDocumentServer wordProcessor)
        {
            #region #AppendTextToRange
            Document document = wordProcessor.Document;
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
        static void AppendToParagraph(RichEditDocumentServer wordProcessor)
        {
            #region #AppendToParagraph
            Document document = wordProcessor.Document;
            document.BeginUpdate();
            document.AppendText("First Paragraph\nSecond Paragraph\nThird Paragraph");
            document.EndUpdate();            
            #endregion #AppendToParagraph
        }
    }
}


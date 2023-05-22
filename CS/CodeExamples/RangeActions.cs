using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class RangeActions
    {
        public static Action<RichEditDocumentServer> InsertTextInRangeAction = InsertTextInRange;
        public static Action<RichEditDocumentServer> AppendTextToRangeAction = AppendTextToRange;
        public static Action<RichEditDocumentServer> AppendToParagraphAction = AppendToParagraph;        

        static void InsertTextInRange(RichEditDocumentServer wordProcessor)
        {
            #region #InsertTextInRange
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("ABCDEFGH");

            // Create the first document range.
            DocumentRange range1 = document.CreateRange(1, 3);

            // Insert text into the first document range
            // and access the range of the inserted text.
            DocumentRange range2 = document.InsertText(range1.End, ">>NewText<<");

            // Output the start and end positions of the first document range. 
            string text1 = String.Format("Range range1 starts at {0}, ends at {1}", range1.Start, range1.End);
            document.Paragraphs.Append();
            document.AppendText(text1);

            // Output the start and end positions of the second document range. 
            string text2 = String.Format("Range range2 starts at {0}, ends at {1}", range2.Start, range2.End);
            document.Paragraphs.Append();
            document.AppendText(text2);
            #endregion #InsertTextInRange
        }

        static void AppendTextToRange(RichEditDocumentServer wordProcessor)
        {
            #region #AppendTextToRange
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("abcdefgh");
            
            // Append text and access the range of the added text.
            DocumentRange r1 = document.AppendText("X");
            string s1 = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End);

            // Append text and access the updated range of the added text.
            document.AppendText("Y");
            document.AppendText("Z");
            string s2 = String.Format("Currently range r1 starts at {0}, ends at {1}", r1.Start, r1.End);

            // Output the start and end positions of the document range. 
            document.Paragraphs.Append();
            document.AppendText(s1);

            // Output the updated start and end positions of the document range.
            document.Paragraphs.Append();
            document.AppendText(s2);
            #endregion #AppendTextToRange
        }
        static void AppendToParagraph(RichEditDocumentServer wordProcessor)
        {
            #region #AppendToParagraph
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the end of each paragraph.
            document.AppendText("First Paragraph\nSecond Paragraph\nThird Paragraph");
            
            // Finalize to edit the document.
            document.EndUpdate();

            // Access the end position of the document range.
            DocumentPosition pos = document.Range.End;

            // Append text to the end of the last paragraph.
            SubDocument doc = pos.BeginUpdateDocument();
            Paragraph par = doc.Paragraphs.Get(pos);
            DocumentPosition newPos = doc.CreatePosition(par.Range.End.ToInt() - 1);
            doc.InsertText(newPos, "<<Appended to Paragraph End>>");
            pos.EndUpdateDocument(doc);
            #endregion #AppendToParagraph
        }
    }
}


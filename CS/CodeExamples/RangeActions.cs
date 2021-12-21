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
        public static Action<RichEditDocumentServer> SelectTextInRangeAction = SelectTextInRange;
        public static Action<RichEditDocumentServer> InsertTextInRangeAction = InsertTextInRange;
        public static Action<RichEditDocumentServer> AppendTextToRangeAction = AppendTextToRange;
        public static Action<RichEditDocumentServer> AppendToParagraphAction = AppendToParagraph;

        static void SelectTextInRange(RichEditDocumentServer wordProcessor)
        {
            #region #SelectTextInRange
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Create a document range.
            DocumentPosition myStart = document.CreatePosition(69);
            DocumentRange myRange = document.CreateRange(myStart, 716);
            
            // Select text in the target range.
            document.Selection = myRange;
            #endregion #SelectTextInRange
        }        

        static void InsertTextInRange(RichEditDocumentServer wordProcessor)
        {
            #region #InsertTextInRange
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("ABCDEFGH");

            // Create the first document range.
            DocumentRange r1 = document.CreateRange(1, 3);

            // Insert text into the first document range
            // and access the range of the inserted text.
            DocumentPosition pos1 = document.CreatePosition(2);
            DocumentRange r2 = document.InsertText(pos1, ">>NewText<<");

            // Output the start and end positions of the first document range. 
            string s1 = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End);
            document.Paragraphs.Append();
            document.AppendText(s1);

            // Output the start and end positions of the second document range. 
            string s2 = String.Format("Range r2 starts at {0}, ends at {1}", r2.Start, r2.End);
            document.Paragraphs.Append();
            document.AppendText(s2);
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


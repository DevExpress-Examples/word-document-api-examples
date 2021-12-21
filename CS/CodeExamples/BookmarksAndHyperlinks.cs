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
    public static class BookmarksAndHyperlinksActions {

    public static Action<RichEditDocumentServer> InsertBookmarkAction = InsertBookmark;
    public static Action<RichEditDocumentServer> InsertHyperlinkAction = InsertHyperlink;
    
         static void InsertBookmark(RichEditDocumentServer wordProcessor)
        {
            #region #InsertBookmark
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
          
            // Access a document.
            Document document = wordProcessor.Document;
            
            // Start to edit the document.
            document.BeginUpdate();

            // Access the start position of the document range.
            DocumentPosition pos = document.Range.Start;
            
            // Create a bookmark at the document top.
            document.Bookmarks.Create(wordProcessor.Document.CreateRange(pos, 0), "Top");

            // Create a hyperlink that navigates to the created bookmark.
            DocumentPosition pos1 = document.CreatePosition((wordProcessor.Document.Range.End).ToInt() + 25);
            document.Hyperlinks.Create(wordProcessor.Document.InsertText(pos1, "get to the top"));
            document.Hyperlinks[0].Anchor = "Top";
        
            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #InsertBookmark
        }
        static void InsertHyperlink(RichEditDocumentServer wordProcessor)
        {
            #region #InsertHyperlink
            // Access a document.
            Document document = wordProcessor.Document;
            
            // Access the start position of the document range.
            DocumentPosition hPos = wordProcessor.Document.Range.Start;

            // Create a hyperlink at the specified position.
            document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"));

            // Specify the URI to which the hyperlink navigates. 
            document.Hyperlinks[0].NavigateUri = "https://devexpress.com";
            
            // Specify the hyperlink tooltip.
            document.Hyperlinks[0].ToolTip = "DevExpress";
            #endregion #InsertHyperlink
        }



    }
}

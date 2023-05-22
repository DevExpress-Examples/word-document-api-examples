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
            
            // Create a bookmark at the start document position.
            document.Bookmarks.Create(document.CreateRange(document.Range.Start, 0), "Top");

            // Create a hyperlink that navigates to the created bookmark.
            wordProcessor.Document.Paragraphs.Append();
            DocumentRange hyperlinkRange = wordProcessor.Document.AppendText("get to the top");
            document.Hyperlinks.Create(hyperlinkRange);
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

            // Create a hyperlink at the specified position.
            DocumentRange hyperlinkRange = document.InsertText(document.Range.Start, "Follow me!");
            document.Hyperlinks.Create(hyperlinkRange);

            // Specify the URI to which the hyperlink navigates. 
            document.Hyperlinks[0].NavigateUri = "https://devexpress.com";
            
            // Specify the hyperlink tooltip.
            document.Hyperlinks[0].ToolTip = "DevExpress";
            #endregion #InsertHyperlink
        }



    }
}

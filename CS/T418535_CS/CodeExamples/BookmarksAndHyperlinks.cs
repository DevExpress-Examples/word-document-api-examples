using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class BookmarksAndHyperlinksActions
    {
         static void InsertBookmark(RichEditDocumentServer server)
        {
            #region #InsertBookmark
            server.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            server.BeginUpdate();
            Document document = server.Document;
            DocumentPosition pos = document.Range.Start;
            document.Bookmarks.Create(server.Document.CreateRange(pos, 0), "Top");
            //Insert the hyperlink anchored to the created bookmark:
            DocumentPosition pos1 = document.CreatePosition((server.Document.Range.End).ToInt() + 25);
            document.Hyperlinks.Create(server.Document.InsertText(pos1, "get to the top"));
            document.Hyperlinks[0].Anchor = "Top";
            server.EndUpdate();
            #endregion #InsertBookmark
        }
         static void InsertHyperlink(RichEditDocumentServer server)
        {
            #region #InsertHyperlink
            Document document = server.Document;
            DocumentPosition hPos = server.Document.Range.Start;
            server.Document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"));
            document.Hyperlinks[0].NavigateUri = "https://www.devexpress.com/Products/NET/Controls/WinForms/Rich_Editor/";
            document.Hyperlinks[0].ToolTip = "WinForms Rich Text Editor";
            #endregion #InsertHyperlink
        }



    }
}

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
         static void InsertBookmark(RichEditDocumentServer wordProcessor)
        {
            #region #InsertBookmark
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            wordProcessor.BeginUpdate();
            Document document = wordProcessor.Document;
            DocumentPosition pos = document.Range.Start;
            document.Bookmarks.Create(wordProcessor.Document.CreateRange(pos, 0), "Top");
            //Insert the hyperlink anchored to the created bookmark:
            DocumentPosition pos1 = document.CreatePosition((wordProcessor.Document.Range.End).ToInt() + 25);
            document.Hyperlinks.Create(wordProcessor.Document.InsertText(pos1, "get to the top"));
            document.Hyperlinks[0].Anchor = "Top";
            wordProcessor.EndUpdate();
            #endregion #InsertBookmark
        }
         static void InsertHyperlink(RichEditDocumentServer wordProcessor)
        {
            #region #InsertHyperlink
            Document document = wordProcessor.Document;
            DocumentPosition hPos = wordProcessor.Document.Range.Start;
            document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"));
            document.Hyperlinks[0].NavigateUri = "https://www.devexpress.com/Products/NET/Controls/WinForms/Rich_Editor/";
            document.Hyperlinks[0].ToolTip = "WinForms Rich Text Editor";
            #endregion #InsertHyperlink
        }



    }
}

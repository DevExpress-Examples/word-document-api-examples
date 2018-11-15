using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using System.Diagnostics;
using DevExpress.XtraRichEdit.Services;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.Export;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
   public static class BasicActions
    {
        static void CreateNewDocument(RichEditDocumentServer server)
        {
            #region #CreateDocument
            server.CreateNewDocument();
            #endregion #CreateDocument
        }
        static void LoadDocument(RichEditDocumentServer server)
        {
            #region #LoadDocument
            server.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            #endregion #LoadDocument
        }
        static void MergeDocuments(RichEditDocumentServer server)
        {
            #region #MergeDocuments
            server.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml);
            server.Document.AppendDocumentContent("Documents//MovieRentals.docx",DocumentFormat.OpenXml);
            #endregion #MergeDocuments
        }
        static void SplitDocument(RichEditDocumentServer server)
        {
            #region #SplitDocument
            server.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            //Split a document per page
            int pageCount = server.DocumentLayout.GetPageCount();
            for (int i = 0; i < pageCount; i++)
            {
                DevExpress.XtraRichEdit.API.Layout.LayoutPage layoutPage = server.DocumentLayout.GetPage(i);
                DevExpress.XtraRichEdit.API.Native.DocumentRange mainBodyRange = server.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length);
                using (RichEditDocumentServer tempServer = new RichEditDocumentServer())
                {
                    tempServer.Document.AppendDocumentContent(mainBodyRange);
                    //Delete last empty paragraph
                    tempServer.Document.Delete(tempServer.Document.Paragraphs.First().Range);
                    //Save the result
                    string fileName = String.Format("doc{0}.rtf", i);
                    tempServer.SaveDocument(fileName, DocumentFormat.Rtf);
                }                
            }
            System.Diagnostics.Process.Start("explorer.exe", "/select," + "doc0.rtf");
            #endregion #SplitDocument
        }
        static void SaveDocument(RichEditDocumentServer server)
        {            
            #region #SaveDocument
            server.Document.AppendDocumentContent("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            server.SaveDocument("SavedDocument.docx", DocumentFormat.OpenXml); 
                System.Diagnostics.Process.Start("explorer.exe", "/select," + "SavedDocument.docx");
            #endregion #SaveDocument
        }
        static void PrintDocument(RichEditDocumentServer server)
        {
            #region #PrintDocument
            server.Document.AppendDocumentContent("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            server.Print();
            #endregion #PrintDocument
        }
    }
}

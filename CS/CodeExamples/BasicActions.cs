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

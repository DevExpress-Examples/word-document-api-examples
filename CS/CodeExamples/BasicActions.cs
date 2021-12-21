using System;
using System.Linq;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
   public static class BasicActions {

        public static Action<RichEditDocumentServer> CreateNewDocumentAction = CreateNewDocument;
        public static Action<RichEditDocumentServer> LoadDocumentAction = LoadDocument;
        public static Action<RichEditDocumentServer> MergeDocumentsAction = MergeDocuments;
        public static Action<RichEditDocumentServer> SplitDocumentAction = SplitDocument;
        public static Action<RichEditDocumentServer> SaveDocumentAction = SaveDocument;
        public static Action<RichEditDocumentServer> PrintDocumentAction = PrintDocument;
  
        static void CreateNewDocument(RichEditDocumentServer wordProcessor)
        {
            #region #CreateDocument
            // Create a new blank document.
            wordProcessor.CreateNewDocument();
            #endregion #CreateDocument
        }
        static void LoadDocument(RichEditDocumentServer wordProcessor)
        {
            #region #LoadDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            #endregion #LoadDocument
        }
        static void MergeDocuments(RichEditDocumentServer wordProcessor)
        {
            #region #MergeDocuments
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml);

            // Insert content from the file at the document end.
            wordProcessor.Document.AppendDocumentContent("Documents//MovieRentals.docx",DocumentFormat.OpenXml);
            #endregion #MergeDocuments
        }
        static void SplitDocument(RichEditDocumentServer wordProcessor)
        {
            #region #SplitDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Obtain a number of pages in the document.
            int pageCount = wordProcessor.DocumentLayout.GetPageCount();
            
            // Check all pages in the document.
            for (int i = 0; i < pageCount; i++)
            {
                // Access the document page.  
                DevExpress.XtraRichEdit.API.Layout.LayoutPage layoutPage = wordProcessor.DocumentLayout.GetPage(i);

                // Access the range of the page's main area.
                DevExpress.XtraRichEdit.API.Native.DocumentRange mainBodyRange = wordProcessor.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length);

                // Create the temporary RichEditDocumentServer instance.
                using (RichEditDocumentServer tempWordProcessor = new RichEditDocumentServer())
                {
                    // Insert the page content to the instance.
                    tempWordProcessor.Document.AppendDocumentContent(mainBodyRange);
                    // Delete the first empty paragraph.
                    tempWordProcessor.Document.Delete(tempWordProcessor.Document.Paragraphs.First().Range);
                    // Save the document page as an RTF file.
                    string fileName = String.Format("doc{0}.rtf", i);
                    tempWordProcessor.SaveDocument(fileName, DocumentFormat.Rtf);
                }                
            }
            // Open the File Explorer and select the saved file.
            System.Diagnostics.Process.Start("explorer.exe", "/select," + "doc0.rtf");
            #endregion #SplitDocument
        }
        static void SaveDocument(RichEditDocumentServer wordProcessor)
        {
            #region #SaveDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Save the document as a DOCX file.
            wordProcessor.SaveDocument("SavedDocument.docx", DocumentFormat.OpenXml);
            
            // Open the File Explorer and select the saved file.
            System.Diagnostics.Process.Start("explorer.exe", "/select," + "SavedDocument.docx");
            #endregion #SaveDocument
        }
        static void PrintDocument(RichEditDocumentServer wordProcessor)
        {
            #region #PrintDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            
            // Print the document to the default printer with the default settings.
            wordProcessor.Print();
            #endregion #PrintDocument
        }
    }
}

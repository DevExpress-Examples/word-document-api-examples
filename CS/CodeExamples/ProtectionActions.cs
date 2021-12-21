using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class ProtectionActions
    {
        public static Action<RichEditDocumentServer> ProtectDocumentAction = ProtectDocument;
        public static Action<RichEditDocumentServer> UnprotectDocumentAction = UnprotectDocument;
        public static Action<RichEditDocumentServer> CreateRangePermissionsAction = CreateRangePermissions;
        static void ProtectDocument(RichEditDocumentServer wordProcessor)
        {
            UnprotectResultingDocument(wordProcessor);
            #region #ProtectDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx",DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Check whether the document is protected.
            if (!document.IsDocumentProtected)
            {
                // Protect the document with a password.
                document.Protect("123", DocumentProtectionType.ReadOnly);
               
                // Create a comment related to the first paragraph.
                document.Comments.Create(document.Paragraphs[0].Range, "Admin");
                
                // Access the comment content.
                SubDocument commentDocument = document.Comments[0].BeginUpdate();                
                
                // Specify the comment text to indicate that the document is protected.
                commentDocument.InsertText(commentDocument.CreatePosition(0), 
                "Document is protected with a password.\nYou cannot modify the document until protection is removed.");
                
                // Finalize to edit the comment.
                commentDocument.EndUpdate();

                // Save and open the protected document.
                wordProcessor.SaveDocument("ResultProtected.docx", DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start("ResultProtected.docx");
            }
            #endregion #ProtectDocument
        }
        static void UnprotectDocument(RichEditDocumentServer wordProcessor)
        {
            #region #UnprotectDocument
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm_Protected.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Check whether the document is protected.
            if (document.IsDocumentProtected == true)
            {
                // Unprotect the document.
                document.Unprotect();

                // Create a comment related to the first paragraph.
                document.Comments.Create(document.Paragraphs[0].Range,"Admin");

                // Access the comment content.
                SubDocument commentDocument = document.Comments[0].BeginUpdate();

                // Specify the comment text to indicate that the document is unprotected.
                commentDocument.InsertText(commentDocument.CreatePosition(0),
               "Document is unprotected. You can modify the document according to your requests.");

                // Finalize to edit the comment.
                commentDocument.EndUpdate();

                // Save and open the protected document.
                wordProcessor.SaveDocument("ResultUnrotected.docx", DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start("ResultUnprotected.docx");
            }
            #endregion #UnprotectDocument
        }
        static void CreateRangePermissions(RichEditDocumentServer wordProcessor)
        {
            UnprotectResultingDocument(wordProcessor);
            #region #CreateRangePermissions
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the range permissions collection.
            RangePermissionCollection rangePermissions = document.BeginUpdateRangePermissions();

            if (document.Paragraphs.Count > 3)
            {
                // Specify the group of users and the user that are allowed to edit the document range.
                RangePermission rp = rangePermissions.CreateRangePermission(document.Paragraphs[3].Range);
                rp.Group = "Administrators";
                rp.UserName = "admin@somecompany.com";
                rangePermissions.Add(rp);
            }

            // Finalize to update the range permissions collection.
            document.EndUpdateRangePermissions(rangePermissions);

            // Protect the document with a password.
            document.Protect("123");

            // Save and open the protected document.
            wordProcessor.SaveDocument("ResultProtected.docx", DocumentFormat.OpenXml);
            System.Diagnostics.Process.Start("ResultProtected.docx");
            #endregion #CreateRangePermissions
        }

        static void UnprotectResultingDocument(RichEditDocumentServer wordProcessor)
        {
            try
            {
                // Load a document from a file.
                wordProcessor.LoadDocument("ResultProtected.docx", DocumentFormat.OpenXml);

                // Access a document.
                Document document = wordProcessor.Document;
                if (document.IsDocumentProtected == true)
                {
                    // Unprotect the document.
                    document.Unprotect();
                }
            }
            catch { }
        }
    }
}

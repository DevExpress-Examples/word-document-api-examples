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
        static void ProtectDocument(RichEditDocumentServer wordProcessor)
        {
            #region #ProtectDocument
            wordProcessor.LoadDocument("Documents//Grimm.docx",DocumentFormat.OpenXml);
            Document document = wordProcessor.Document;
            if (!document.IsDocumentProtected)
            {
                //Protect the document with a password
                document.Protect("123",DocumentProtectionType.ReadOnly);

                //Insert a comment indicating that the document is protected
                document.Comments.Create(document.Paragraphs[0].Range, "Admin");                
                SubDocument commentDocument = document.Comments[0].BeginUpdate();                
                commentDocument.InsertText(commentDocument.CreatePosition(0), 
                "Document is protected with a password.\nYou cannot modify the document until protection is removed.");
                commentDocument.EndUpdate();
            }
            #endregion #ProtectDocument
        }
        static void UnprotectDocument(RichEditDocumentServer wordProcessor)
        {
            #region #UnprotectDocument
            wordProcessor.LoadDocument("Documents//Grimm_Protected.docx", DocumentFormat.OpenXml);
            Document document = wordProcessor.Document;

            if (document.IsDocumentProtected == true)
            {
                //Unprotect the document
                document.Unprotect();
                
                //Insert a comment indicating that the document can be edited
                document.Comments.Create(document.Paragraphs[0].Range,"Admin");
                SubDocument commentDocument = document.Comments[0].BeginUpdate();
                commentDocument.InsertText(commentDocument.CreatePosition(0),
               "Document is unprotected. You can modify the document according to your requests.");
                commentDocument.EndUpdate();
            }
            #endregion #UnprotectDocument
        }
        static void CreateRangePermissions(RichEditDocumentServer wordProcessor)
        {
            #region #CreateRangePermissions
            wordProcessor.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml);
            Document document = wordProcessor.Document;

            // Protect document range
            RangePermissionCollection rangePermissions = document.BeginUpdateRangePermissions();
            RangePermission rp = rangePermissions.CreateRangePermission(document.Paragraphs[3].Range);
            rp.Group = "Administrators";
            rp.UserName = "admin@somecompany.com";
            rangePermissions.Add(rp);

            document.EndUpdateRangePermissions(rangePermissions);
            // Enforce protection and set password.
            document.Protect("123");
            #endregion #CreateRangePermissions
        }
    }
}

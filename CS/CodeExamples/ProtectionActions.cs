﻿using System;
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
        static void ProtectDocument(RichEditDocumentServer server)
        {
            #region #ProtectDocument
            server.LoadDocument("Documents//Grimm.docx",DocumentFormat.OpenXml);
            Document document = server.Document;
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
        static void UnprotectDocument(RichEditDocumentServer server)
        {
            #region #UnprotectDocument
            server.LoadDocument("Documents//Grimm_Protected.docx", DocumentFormat.OpenXml);
            Document document = server.Document;

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
        static void CreateRangePermissions(RichEditDocumentServer server)
        {
            #region #CreateRangePermissions
            server.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml);
            Document document = server.Document;

            // Protect document range
            RangePermissionCollection rangePermissions = document.BeginUpdateRangePermissions();
            RangePermission rp = rangePermissions.CreateRangePermission(document.Paragraphs[3].Range);
            rp.Group = "Everyone";
            rangePermissions.Add(rp);

            document.EndUpdateRangePermissions(rangePermissions);
            // Enforce protection and set password.
            document.Protect("123");
            #endregion #CreateRangePermissions
        }
    }
}

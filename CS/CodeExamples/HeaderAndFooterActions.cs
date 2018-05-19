﻿using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditEDocumentServerExample.CodeExamples
{
    class HeadersAndFootersActions
    {
       
        static void CreateHeader(RichEditDocumentServer server)
        {
            #region #CreateHeader
            Document document = server.Document;
            Section firstSection = document.Sections[0];
            // Create an empty header.
            SubDocument newHeader = firstSection.BeginUpdateHeader();
            firstSection.EndUpdateHeader(newHeader);
            // Check whether the document already has a header (the same header for all pages).
            if (firstSection.HasHeader(HeaderFooterType.Primary))
            {
                SubDocument headerDocument = firstSection.BeginUpdateHeader();
                document.ChangeActiveDocument(headerDocument);
                document.CaretPosition = headerDocument.CreatePosition(0);
                firstSection.EndUpdateHeader(headerDocument);
            }
            #endregion #CreateHeader
        }


        static void ModifyHeader(RichEditDocumentServer server)
        {
            #region #ModifyHeader
            Document document = server.Document;
            document.AppendSection();
            Section firstSection = document.Sections[0];
            // Modify the header of the HeaderFooterType.First type.
            SubDocument myHeader = firstSection.BeginUpdateHeader(HeaderFooterType.First);
            DocumentRange range = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ");
            Field fld = myHeader.Fields.Create(range.End, "PAGE \\* ARABICDASH");
            myHeader.Fields.Update();
            firstSection.EndUpdateHeader(myHeader);
            // Display the header of the HeaderFooterType.First type on the first page.
            firstSection.DifferentFirstPage = true;
            #endregion #ModifyHeader
        }
    }
}

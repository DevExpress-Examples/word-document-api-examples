using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;
using System;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class HeadersAndFootersActions
    {
        public static Action<RichEditDocumentServer> CreateHeaderAction = CreateHeader;
        public static Action<RichEditDocumentServer> ModifyHeaderAction = ModifyHeader;

        static void CreateHeader(RichEditDocumentServer wordProcessor)
        {
            #region #CreateHeader
            // Access a document.
            Document document = wordProcessor.Document;

            // Access the first document section.
            Section firstSection = document.Sections[0];

            // Check whether the document already has a header (the same header for all pages).
            if (!firstSection.HasHeader(HeaderFooterType.Primary))
            {
                // Create a header.
                SubDocument newHeader = firstSection.BeginUpdateHeader();
                newHeader.AppendText("Header");
                firstSection.EndUpdateHeader(newHeader);
            }
            #endregion #CreateHeader
        }


        static void ModifyHeader(RichEditDocumentServer wordProcessor)
        {
            #region #ModifyHeader
            // Access a document.
            Document document = wordProcessor.Document;

            // Append a new section to the document.
            document.AppendSection();

            // Access the first document section.
            Section firstSection = document.Sections[0];

            // Start to edit the header of the HeaderFooterType.First type.
            SubDocument myHeader = firstSection.BeginUpdateHeader(HeaderFooterType.First);
            
            // Change the header text.
            DocumentRange range = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ");
            Field fld = myHeader.Fields.Create(range.End, "PAGE \\* ARABICDASH");
            
            // Update all fields in the header.
            myHeader.Fields.Update();

            // Finalize to edit the header.
            firstSection.EndUpdateHeader(myHeader);

            // Display the header on the first document page.
            firstSection.DifferentFirstPage = true;
            #endregion #ModifyHeader
        }
    }
}

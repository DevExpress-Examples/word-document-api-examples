using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditEDocumentServerExample.CodeExamples
{
    class HeadersAndFootersActions
    {
       
        static void CreateHeader(RichEditDocumentServer wordProcessor)
        {
            #region #CreateHeader
            Document document = wordProcessor.Document;
            Section firstSection = document.Sections[0];

            // Check whether the document already has a header (the same header for all pages).
            if (!firstSection.HasHeader(HeaderFooterType.Primary))
            {
                SubDocument newHeader = firstSection.BeginUpdateHeader();
                newHeader.AppendText("Header");
                firstSection.EndUpdateHeader(newHeader);
            }
            #endregion #CreateHeader
        }


        static void ModifyHeader(RichEditDocumentServer wordProcessor)
        {
            #region #ModifyHeader
            Document document = wordProcessor.Document;
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

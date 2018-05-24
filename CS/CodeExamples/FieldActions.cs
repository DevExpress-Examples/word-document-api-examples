using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class FieldActions
    {
        static void InsertField(RichEditDocumentServer server)
        {
            #region #InsertField
            Document document = server.Document;
            document.BeginUpdate();
            document.Fields.Create(document.CaretPosition, "DATE");
            document.Fields.Update();
            document.EndUpdate();
            #endregion #InsertField
        }

        static void ModifyFieldCode(RichEditDocumentServer server)
        {
            #region #ModifyFieldCode
            Document document = server.Document;
            document.BeginUpdate();
            document.Fields.Create(document.CaretPosition, "DATE");
            document.EndUpdate();
            for (int i = 0; i < document.Fields.Count; i++)
            {
                string fieldCode = document.GetText(document.Fields[i].CodeRange);
                if (fieldCode == "DATE")
                {
                    DocumentPosition position = document.Fields[i].CodeRange.End;
                    document.InsertText(position, @" \@ ""M / d / yyyy HH: mm:ss""");
                }
            }
            document.Fields.Update();
            #endregion #ModifyFieldCode
        }

        static void CreateFieldFromRange(RichEditDocumentServer server)
        {
            #region #CreateFieldFromRange
            Document document = server.Document;
            document.BeginUpdate();
            document.AppendText("SYMBOL 0x54 \\f Wingdings \\s 24");
            document.EndUpdate();
            document.Fields.Create(document.Paragraphs[0].Range);
            document.Fields.Update();
            #endregion #CreateFieldFromRange
        }

        static void ShowFieldCodes(RichEditDocumentServer server)
        {
            #region #ShowFieldCodes
            Document document = server.Document;
            document.LoadDocument("MailMergeSimple.docx", DocumentFormat.OpenXml);
            for (int i = 0; i < document.Fields.Count; i++)
            {
                document.Fields[i].ShowCodes = true;
            }
            #endregion #ShowFieldCodes
        }
    }
}

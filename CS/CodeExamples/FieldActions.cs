using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class FieldActions
    {
        public static Action<RichEditDocumentServer> InsertFieldAction = InsertField;
        public static Action<RichEditDocumentServer> ModifyFieldCodeAction = ModifyFieldCode;
        public static Action<RichEditDocumentServer> CreateFieldFromRangeAction = CreateFieldFromRange;

        static void InsertField(RichEditDocumentServer wordProcessor)
        {
            #region #InsertField
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create the "DATE" field.
            document.Fields.Create(document.Range.Start, "DATE");

            // Update all fields in the main document body.
            document.Fields.Update();

            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #InsertField
        }

        static void ModifyFieldCode(RichEditDocumentServer wordProcessor)
        {
            #region #ModifyFieldCode
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create the "DATE" field.
            document.Fields.Create(document.CaretPosition, "DATE");

            // Finalize to edit the document.
            document.EndUpdate();

            // Check all fields in the document.
            for (int i = 0; i < document.Fields.Count; i++)
            {
                // Access a field code.
                string fieldCode = document.GetText(document.Fields[i].CodeRange);

                // Check whether a field code is "DATE".
                if (fieldCode == "DATE")
                {
                    // Set the document position to the end of the field code range.
                    DocumentPosition position = document.Fields[i].CodeRange.End;
                    // Specify the date and time format for the field. 
                    document.InsertText(position, @" \@ ""M / d / yyyy HH: mm:ss""");
                }
            }
            // Update all fields in the main document body.
            document.Fields.Update();
            #endregion #ModifyFieldCode
        }

        static void CreateFieldFromRange(RichEditDocumentServer wordProcessor)
        {
            #region #CreateFieldFromRange
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Append text to the document.
            document.AppendText("SYMBOL 0x54 \\f Wingdings \\s 24");

            // Finalize to edit the document.
            document.EndUpdate();

            // Convert inserted text to a field.
            document.Fields.Create(document.Paragraphs[0].Range);
            
            // Update all fields in the main document body.
            document.Fields.Update();
            #endregion #CreateFieldFromRange
        }
    }
}

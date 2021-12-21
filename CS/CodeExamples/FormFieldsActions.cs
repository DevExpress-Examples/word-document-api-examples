using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class FormFieldsActions
    {
        public static Action<RichEditDocumentServer> InsertCheckBoxAction = InsertCheckBox;

        static void InsertCheckBox(RichEditDocumentServer wordProcessor)
        {
            #region #InsertCheckbox
            // Access the start position of the document range.
            DocumentPosition currentPosition = wordProcessor.Document.Range.Start;

            // Insert a checkbox at the specified position.
            DevExpress.XtraRichEdit.API.Native.CheckBox checkBox = wordProcessor.Document.FormFields.InsertCheckBox(currentPosition);
            
            // Specify the checkbox properties.
            checkBox.Name = "check1";
            checkBox.State = CheckBoxState.Checked;
            checkBox.SizeMode = CheckBoxSizeMode.Auto;
            checkBox.HelpTextType = FormFieldTextType.Custom;
            checkBox.HelpText = "help text";
            #endregion #InsertCheckbox
        }
    }
}

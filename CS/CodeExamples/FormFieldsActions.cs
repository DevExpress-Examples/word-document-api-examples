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
        static void InsertCheckBox(RichEditDocumentServer wordProcessor)
        {
            #region #InsertCheckbox
            DocumentPosition currentPosition = wordProcessor.Document.Range.Start;
            DevExpress.XtraRichEdit.API.Native.CheckBox checkBox = wordProcessor.Document.FormFields.InsertCheckBox(currentPosition);
            checkBox.Name = "check1";
            checkBox.State = CheckBoxState.Checked;
            checkBox.SizeMode = CheckBoxSizeMode.Auto;
            checkBox.HelpTextType = FormFieldTextType.Custom;
            checkBox.HelpText = "help text";
            #endregion #InsertCheckbox
        }
    }
}

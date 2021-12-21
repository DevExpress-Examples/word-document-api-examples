Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FormFieldsActions

        Private Shared Sub InsertCheckBox(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertCheckbox"
            Dim currentPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = wordProcessor.Document.Range.Start
            Dim checkBox As DevExpress.XtraRichEdit.API.Native.CheckBox = wordProcessor.Document.FormFields.InsertCheckBox(currentPosition)
            checkBox.Name = "check1"
            checkBox.State = DevExpress.XtraRichEdit.API.Native.CheckBoxState.Checked
            checkBox.SizeMode = DevExpress.XtraRichEdit.API.Native.CheckBoxSizeMode.Auto
            checkBox.HelpTextType = DevExpress.XtraRichEdit.API.Native.FormFieldTextType.Custom
            checkBox.HelpText = "help text"
#End Region  ' #InsertCheckbox
        End Sub
    End Class
End Namespace

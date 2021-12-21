Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FormFieldsActions

        Public Shared InsertCheckBoxAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormFieldsActions.InsertCheckBox

        Private Shared Sub InsertCheckBox(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertCheckbox"
            ' Access the start position of the document range.
            Dim currentPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = wordProcessor.Document.Range.Start
            ' Insert a checkbox at the specified position.
            Dim checkBox As DevExpress.XtraRichEdit.API.Native.CheckBox = wordProcessor.Document.FormFields.InsertCheckBox(currentPosition)
            ' Specify the checkbox properties.
            checkBox.Name = "check1"
            checkBox.State = DevExpress.XtraRichEdit.API.Native.CheckBoxState.Checked
            checkBox.SizeMode = DevExpress.XtraRichEdit.API.Native.CheckBoxSizeMode.Auto
            checkBox.HelpTextType = DevExpress.XtraRichEdit.API.Native.FormFieldTextType.Custom
            checkBox.HelpText = "help text"
#End Region  ' #InsertCheckbox
        End Sub
    End Class
End Namespace

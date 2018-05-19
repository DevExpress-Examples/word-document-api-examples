Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class FormFieldsActions
        Private Shared Sub InsertCheckBox(ByVal server As RichEditDocumentServer)
'            #Region "#InsertCheckbox"
            Dim currentPosition As DocumentPosition = server.Document.CaretPosition
            Dim checkBox As DevExpress.XtraRichEdit.API.Native.CheckBox = server.Document.FormFields.InsertCheckBox(currentPosition)
            checkBox.Name = "check1"
            checkBox.State = CheckBoxState.Checked
            checkBox.SizeMode = CheckBoxSizeMode.Auto
            checkBox.HelpTextType = FormFieldTextType.Custom
            checkBox.HelpText = "help text"
'            #End Region ' #InsertCheckbox
        End Sub
    End Class
End Namespace

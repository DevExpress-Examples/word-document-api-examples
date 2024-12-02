Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Drawing

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module ContentControlsActions

        Public CreateContentControlsAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf CreateContentControls

        Public ChangeContentControlsAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ChangeContentControls

        Public RemoveContentControlsAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RemoveContentControls

        Private Sub CreateContentControls(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateContentControls"
            wordProcessor.LoadDocument("Documents\Simple Form.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim contentControls = document.ContentControls
            ' Insert a form to enter a name:
            Dim namePosition = document.CreatePosition(document.Paragraphs(0).Range.[End].ToInt() - 1)
            Dim nameControl = contentControls.InsertPlainTextControl(namePosition)
            ' Insert text in a content control:
            Dim nameTextPosition = document.CreatePosition(nameControl.Range.Start.ToInt() + 1)
            document.InsertText(nameTextPosition, "Click to enter a name")
            ' Insert a drop-down list to select the appointment type:
            Dim listPosition = document.CreatePosition(document.Paragraphs(1).Range.[End].ToInt() - 1)
            Dim listControl = contentControls.InsertDropDownListControl(listPosition)
            ' Add items to the drop-down list:
            listControl.AddItem("First Appointment", "First Appointment")
            listControl.AddItem("Follow-Up Appointment", "Follow-Up Appointment")
            listControl.AddItem("Laboratory Results Check", "Laboratory Results Check")
            listControl.SelectedItemIndex = 1
            ' Insert a date picker to select the appointment date:
            Dim datePosition = document.CreatePosition(document.Paragraphs(2).Range.[End].ToInt() - 1)
            Dim datePicker = contentControls.InsertDatePickerControl(datePosition)
            datePicker.DateFormat = "dddd, MMMM dd, yyyy"
            ' Insert a checkbox:
            Dim checkboxControl = contentControls.InsertCheckboxControl(document.Paragraphs(3).Range.Start)
            checkboxControl.Checked = False
#End Region  ' #CreateContentControls
        End Sub

        Private Sub ChangeContentControls(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeContentControlParameters"
            wordProcessor.LoadDocument("Documents\Simple Form Filled.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim contentControls = document.ContentControls
            For Each contentControl In contentControls
                contentControl.Color = Color.Red
                Select Case contentControl.ControlType
                    Case DevExpress.XtraRichEdit.API.Native.ContentControlType.RichText, DevExpress.XtraRichEdit.API.Native.ContentControlType.PlainText
                        contentControl.IsTemporary = True
                    Case DevExpress.XtraRichEdit.API.Native.ContentControlType.Checkbox
                        Dim checkbox As DevExpress.XtraRichEdit.API.Native.ContentControlCheckbox = TryCast(contentControl, DevExpress.XtraRichEdit.API.Native.ContentControlCheckbox)
                        checkbox.CheckedSymbolStyle.Character = "*"c
                End Select
            Next
#End Region  ' #ChangeContentControlParameters        
        End Sub

        Private Sub RemoveContentControls(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#RemoveContentControls"
            wordProcessor.LoadDocument("Documents\Simple Form Filled.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim contentControls = document.ContentControls
            For i = 0 To contentControls.Count - 1
                If contentControls(i).ControlType Is DevExpress.XtraRichEdit.API.Native.ContentControlType.[Date] Then
                    contentControls.Remove(contentControls(i), True)
                End If
            Next
#End Region  ' #RemoveContentControls
        End Sub
    End Module
End Namespace

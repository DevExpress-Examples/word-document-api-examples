Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FieldActions

        Public Shared InsertFieldAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FieldActions.InsertField

        Public Shared ModifyFieldCodeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FieldActions.ModifyFieldCode

        Public Shared CreateFieldFromRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FieldActions.CreateFieldFromRange

        Public Shared ShowFieldCodesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FieldActions.ShowFieldCodes

        Private Shared Sub InsertField(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertField"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create the "DATE" field.
            document.Fields.Create(document.Range.Start, "DATE")
            ' Update all fields in the main document body.
            document.Fields.Update()
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #InsertField
        End Sub

        Private Shared Sub ModifyFieldCode(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ModifyFieldCode"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create the "DATE" field.
            document.Fields.Create(document.CaretPosition, "DATE")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Check all fields in the document.
            For i As Integer = 0 To document.Fields.Count - 1
                ' Access a field code.
                Dim fieldCode As String = document.GetText(document.Fields(CInt((i))).CodeRange)
                ' Check whether a field code is "DATE".
                If Equals(fieldCode, "DATE") Then
                    ' Set the document position to the end of the field code range.
                    Dim position As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Fields(CInt((i))).CodeRange.[End]
                    ' Specify the date and time format for the field. 
                    document.InsertText(position, " \@ ""M / d / yyyy HH: mm:ss""")
                End If
            Next

            ' Update all fields in the main document body.
            document.Fields.Update()
#End Region  ' #ModifyFieldCode
        End Sub

        Private Shared Sub CreateFieldFromRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateFieldFromRange"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText("SYMBOL 0x54 \f Wingdings \s 24")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Convert inserted text to a field.
            document.Fields.Create(document.Paragraphs(CInt((0))).Range)
            ' Update all fields in the main document body.
            document.Fields.Update()
#End Region  ' #CreateFieldFromRange
        End Sub

        Private Shared Sub ShowFieldCodes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ShowFieldCodes"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\MailMergeSimple.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check all fields in the main document body.
            For i As Integer = 0 To document.Fields.Count - 1
                ' Show field codes.
                document.Fields(CInt((i))).ShowCodes = True
            Next
#End Region  ' #ShowFieldCodes
        End Sub
    End Class
End Namespace

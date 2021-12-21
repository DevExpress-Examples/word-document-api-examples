Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FieldActions

        Private Shared Sub InsertField(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertField"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.Fields.Create(document.Range.Start, "DATE")
            document.Fields.Update()
            document.EndUpdate()
#End Region  ' #InsertField
        End Sub

        Private Shared Sub ModifyFieldCode(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ModifyFieldCode"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.Fields.Create(document.CaretPosition, "DATE")
            document.EndUpdate()
            For i As Integer = 0 To document.Fields.Count - 1
                Dim fieldCode As String = document.GetText(document.Fields(CInt((i))).CodeRange)
                If Equals(fieldCode, "DATE") Then
                    Dim position As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Fields(CInt((i))).CodeRange.[End]
                    document.InsertText(position, " \@ ""M / d / yyyy HH: mm:ss""")
                End If
            Next

            document.Fields.Update()
#End Region  ' #ModifyFieldCode
        End Sub

        Private Shared Sub CreateFieldFromRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateFieldFromRange"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.AppendText("SYMBOL 0x54 \f Wingdings \s 24")
            document.EndUpdate()
            document.Fields.Create(document.Paragraphs(CInt((0))).Range)
            document.Fields.Update()
#End Region  ' #CreateFieldFromRange
        End Sub

        Private Shared Sub ShowFieldCodes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ShowFieldCodes"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("MailMergeSimple.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            For i As Integer = 0 To document.Fields.Count - 1
                document.Fields(CInt((i))).ShowCodes = True
            Next
#End Region  ' #ShowFieldCodes
        End Sub
    End Class
End Namespace

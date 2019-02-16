Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class FieldActions
        Private Shared Sub InsertField(ByVal server As RichEditDocumentServer)
'            #Region "#InsertField"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.Fields.Create(document.CaretPosition, "DATE")
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #InsertField
        End Sub

        Private Shared Sub ModifyFieldCode(ByVal server As RichEditDocumentServer)
'            #Region "#ModifyFieldCode"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.Fields.Create(document.CaretPosition, "DATE")
            document.EndUpdate()
            For i As Integer = 0 To document.Fields.Count - 1
                Dim fieldCode As String = document.GetText(document.Fields(i).CodeRange)
                If fieldCode = "DATE" Then
                    Dim position As DocumentPosition = document.Fields(i).CodeRange.End
                    document.InsertText(position, " \@ ""M / d / yyyy HH: mm:ss""")
                End If
            Next i
            document.Fields.Update()
'            #End Region ' #ModifyFieldCode
        End Sub

        Private Shared Sub CreateFieldFromRange(ByVal server As RichEditDocumentServer)
'            #Region "#CreateFieldFromRange"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.AppendText("SYMBOL 0x54 \f Wingdings \s 24")
            document.EndUpdate()
            document.Fields.Create(document.Paragraphs(0).Range)
            document.Fields.Update()
'            #End Region ' #CreateFieldFromRange
        End Sub

        Private Shared Sub ShowFieldCodes(ByVal server As RichEditDocumentServer)
'            #Region "#ShowFieldCodes"
            Dim document As Document = server.Document
            document.LoadDocument("MailMergeSimple.docx", DocumentFormat.OpenXml)
            For i As Integer = 0 To document.Fields.Count - 1
                document.Fields(i).ShowCodes = True
            Next i
'            #End Region ' #ShowFieldCodes
        End Sub
    End Class
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class RangeActions

        Private Shared Sub SelectTextInRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SelectTextInRange"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim myStart As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(69)
            Dim myRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(myStart, 216)
            document.Selection = myRange
#End Region  ' #SelectTextInRange
        End Sub

        Private Shared Sub InsertTextInRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertTextInRange"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.AppendText("ABCDEFGH")
            Dim r1 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(1, 3)
            Dim pos1 As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(2)
            Dim r2 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.InsertText(pos1, ">>NewText<<")
            Dim s1 As String = System.[String].Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            Dim s2 As String = System.[String].Format("Range r2 starts at {0}, ends at {1}", r2.Start, r2.[End])
            document.Paragraphs.Append()
            document.AppendText(s1)
            document.Paragraphs.Append()
            document.AppendText(s2)
#End Region  ' #InsertTextInRange
        End Sub

        Private Shared Sub AppendTextToRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AppendTextToRange"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.AppendText("abcdefgh")
            Dim r1 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.AppendText("X")
            Dim s1 As String = System.[String].Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            document.AppendText("Y")
            document.AppendText("Z")
            Dim s2 As String = System.[String].Format("Currently range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            document.Paragraphs.Append()
            document.AppendText(s1)
            document.Paragraphs.Append()
            document.AppendText(s2)
#End Region  ' #AppendTextToRange
        End Sub

        Private Shared Sub AppendToParagraph(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AppendToParagraph"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.AppendText("First Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Second Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Third Paragraph")
            document.EndUpdate()
#End Region  ' #AppendToParagraph
        End Sub
    End Class
End Namespace

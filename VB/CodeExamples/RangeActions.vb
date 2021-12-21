Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class RangeActions

        Public Shared SelectTextInRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.SelectTextInRange

        Public Shared InsertTextInRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.InsertTextInRange

        Public Shared AppendTextToRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.AppendTextToRange

        Public Shared AppendToParagraphAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.AppendToParagraph

        Private Shared Sub SelectTextInRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SelectTextInRange"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a document range.
            Dim myStart As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(69)
            Dim myRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(myStart, 716)
            ' Select text in the target range.
            document.Selection = myRange
#End Region  ' #SelectTextInRange
        End Sub

        Private Shared Sub InsertTextInRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertTextInRange"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Append text to the document.
            document.AppendText("ABCDEFGH")
            ' Create the first document range.
            Dim r1 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(1, 3)
            ' Insert text into the first document range
            ' and access the range of the inserted text.
            Dim pos1 As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(2)
            Dim r2 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.InsertText(pos1, ">>NewText<<")
            ' Output the start and end positions of the first document range. 
            Dim s1 As String = System.[String].Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            document.Paragraphs.Append()
            document.AppendText(s1)
            ' Output the start and end positions of the second document range. 
            Dim s2 As String = System.[String].Format("Range r2 starts at {0}, ends at {1}", r2.Start, r2.[End])
            document.Paragraphs.Append()
            document.AppendText(s2)
#End Region  ' #InsertTextInRange
        End Sub

        Private Shared Sub AppendTextToRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AppendTextToRange"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Append text to the document.
            document.AppendText("abcdefgh")
            ' Append text and access the range of the added text.
            Dim r1 As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.AppendText("X")
            Dim s1 As String = System.[String].Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            ' Append text and access the updated range of the added text.
            document.AppendText("Y")
            document.AppendText("Z")
            Dim s2 As String = System.[String].Format("Currently range r1 starts at {0}, ends at {1}", r1.Start, r1.[End])
            ' Output the start and end positions of the document range. 
            document.Paragraphs.Append()
            document.AppendText(s1)
            ' Output the updated start and end positions of the document range.
            document.Paragraphs.Append()
            document.AppendText(s2)
#End Region  ' #AppendTextToRange
        End Sub

        Private Shared Sub AppendToParagraph(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AppendToParagraph"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the end of each paragraph.
            document.AppendText("First Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Second Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Third Paragraph")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Access the end position of the document range.
            Dim pos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Range.[End]
            ' Append text to the end of the last paragraph.
            Dim doc As DevExpress.XtraRichEdit.API.Native.SubDocument = pos.BeginUpdateDocument()
            Dim par As DevExpress.XtraRichEdit.API.Native.Paragraph = doc.Paragraphs.[Get](pos)
            Dim newPos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = doc.CreatePosition(par.Range.[End].ToInt() - 1)
            doc.InsertText(newPos, "<<Appended to Paragraph End>>")
            pos.EndUpdateDocument(doc)
#End Region  ' #AppendToParagraph
        End Sub
    End Class
End Namespace

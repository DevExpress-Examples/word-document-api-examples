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

        Public Shared InsertTextInRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.InsertTextInRange

        Public Shared AppendTextToRangeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.AppendTextToRange

        Public Shared AppendToParagraphAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.RangeActions.AppendToParagraph

        Private Shared Sub InsertTextInRange(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertTextInRange"
            ' Access a document.
            Dim document As Document = wordProcessor.Document

            ' Append text to the document.
            document.AppendText("ABCDEFGH")

            ' Create the first document range.
            Dim range1 As DocumentRange = document.CreateRange(1, 3)

            ' Insert text into the first document range
            ' and access the range of the inserted text.
            Dim range2 As DocumentRange = document.InsertText(range1.End, ">>NewText<<")

            ' Output the start and end positions of the first document range. 
            Dim text1 As String = String.Format("Range range1 starts at {0}, ends at {1}", range1.Start, range1.[End])
            document.Paragraphs.Append()
            document.AppendText(text1)

            ' Output the start and end positions of the second document range. 
            Dim text2 As String = String.Format("Range range2 starts at {0}, ends at {1}", range2.Start, range2.[End])
            document.Paragraphs.Append()
            document.AppendText(text2)
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

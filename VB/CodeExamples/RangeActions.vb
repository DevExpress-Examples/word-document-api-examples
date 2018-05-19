Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class RangeActions
        Private Shared Sub SelectTextInRange(ByVal server As RichEditDocumentServer)
'            #Region "#SelectTextInRange"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            Dim myStart As DocumentPosition = document.CreatePosition(69)
            Dim myRange As DocumentRange = document.CreateRange(myStart, 216)
            document.Selection = myRange
'            #End Region ' #SelectTextInRange
        End Sub

        Private Shared Sub InsertTextAtCaretPosition(ByVal server As RichEditDocumentServer)
'            #Region "#InsertTextAtCaretPosition"
            Dim document As Document = server.Document
            Dim pos As DocumentPosition = document.CaretPosition
            Dim doc As SubDocument = pos.BeginUpdateDocument()
            doc.InsertText(pos, " INSERTED TEXT ")
            pos.EndUpdateDocument(doc)
'            #End Region ' #InsertTextAtCaretPosition
        End Sub

        Private Shared Sub InsertTextInRange(ByVal server As RichEditDocumentServer)
'            #Region "#InsertTextInRange"
            Dim document As Document = server.Document
            document.AppendText("ABCDEFGH")
            Dim r1 As DocumentRange = document.CreateRange(1, 3)
            Dim pos1 As DocumentPosition = document.CreatePosition(2)
            Dim r2 As DocumentRange = document.InsertText(pos1, ">>NewText<<")
            Dim s1 As String = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End)
            Dim s2 As String = String.Format("Range r2 starts at {0}, ends at {1}", r2.Start, r2.End)
            document.Paragraphs.Append()
            document.AppendText(s1)
            document.Paragraphs.Append()
            document.AppendText(s2)
'            #End Region ' #InsertTextInRange
        End Sub

        Private Shared Sub AppendTextToRange(ByVal server As RichEditDocumentServer)
'            #Region "#AppendTextToRange"
            Dim document As Document = server.Document
            document.AppendText("abcdefgh")
            Dim r1 As DocumentRange = document.AppendText("X")
            Dim s1 As String = String.Format("Range r1 starts at {0}, ends at {1}", r1.Start, r1.End)
            document.AppendText("Y")
            document.AppendText("Z")
            Dim s2 As String = String.Format("Currently range r1 starts at {0}, ends at {1}", r1.Start, r1.End)
            document.Paragraphs.Append()
            document.AppendText(s1)
            document.Paragraphs.Append()
            document.AppendText(s2)
'            #End Region ' #AppendTextToRange
        End Sub

        Private Shared Sub CopyAndPasteRange(ByVal server As RichEditDocumentServer)
'            #Region "#CopyAndPasteRange"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            Dim myRange As DocumentRange = document.Paragraphs(0).Range
            document.Copy(myRange)
            document.Paste(DocumentFormat.PlainText)
'            #End Region ' #CopyAndPasteRange
        End Sub

        Private Shared Sub AppendToParagraph(ByVal server As RichEditDocumentServer)
'            #Region "#AppendToParagraph"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.AppendText("First Paragraph" & vbLf & "Second Paragraph" & vbLf & "Third Paragraph")
            document.EndUpdate()
            Dim pos As DocumentPosition = document.CaretPosition
            Dim doc As SubDocument = pos.BeginUpdateDocument()
            Dim par As Paragraph = doc.Paragraphs.Get(pos)
            Dim newPos As DocumentPosition = doc.CreatePosition(par.Range.End.ToInt() - 1)
            doc.InsertText(newPos, "<<Appended to Paragraph End>>")
            pos.EndUpdateDocument(doc)
'            #End Region ' #AppendToParagraph
        End Sub
    End Class
End Namespace


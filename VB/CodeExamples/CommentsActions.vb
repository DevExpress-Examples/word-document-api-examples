Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
Imports System.Diagnostics

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Class CommentsActions

        Public Shared CreateCommentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.CommentsActions.CreateComment

        Public Shared CreateNestedCommentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.CommentsActions.CreateNestedComment

        Public Shared DeleteCommentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.CommentsActions.DeleteComment

        Public Shared EditCommentPropertiesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.CommentsActions.EditCommentProperties

        Public Shared EditCommentContentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.CommentsActions.EditCommentContent

        Private Shared Sub CreateComment(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateComment"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Paragraphs.Count > 2 Then
                ' Access the range of the third paragraph.
                Dim docRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((2))).Range
                ' Specify the comment's author name.
                Dim commentAuthor As String = "Johnson Alphonso D"
                ' Create a comment.
                document.Comments.Create(docRange, commentAuthor, System.DateTime.Now)
            End If
#End Region  ' #CreateComment
        End Sub

        Private Shared Sub CreateNestedComment(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNestedComment"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Comments.Count > 1 Then
                ' Create a new comment nested in the parent comment.
                Dim newComment As DevExpress.XtraRichEdit.API.Native.Comment = document.Comments.Create("Vicars Anny", document.Comments(1))
                newComment.[Date] = System.DateTime.Now
                Dim commentDocument As SubDocument = newComment.BeginUpdate()
                commentDocument.InsertText(commentDocument.Range.Start, "I agree")
                newComment.EndUpdate(commentDocument)
            End If
#End Region  ' #CreateNestedComment
        End Sub

        Private Shared Sub DeleteComment(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#DeleteComment"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Comments.Count > 0 Then
                ' Delete the first comment.
                document.Comments.Remove(document.Comments(0))
            End If
#End Region  ' #DeleteComment
        End Sub

        Private Shared Sub EditCommentProperties(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditCommentProperties"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim commentCount As Integer = document.Comments.Count
            If commentCount > 0 Then
                ' Start to edit the document.
                document.BeginUpdate()
                ' Access a comment and edit its properties.
                Dim comment As DevExpress.XtraRichEdit.API.Native.Comment = document.Comments(document.Comments.Count - 1)
                comment.Name = "New Name"
                comment.[Date] = System.DateTime.Now
                comment.Author = "New Author"
                ' Finalize to edit the document.
                document.EndUpdate()
            End If
#End Region  ' #EditCommentProperties
        End Sub

        Private Shared Sub EditCommentContent(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditCommentContent"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim commentCount As Integer = document.Comments.Count
            If commentCount > 0 Then
                ' Access a comment.
                Dim comment As DevExpress.XtraRichEdit.API.Native.Comment = document.Comments(document.Comments.Count - 1)
                If comment IsNot Nothing Then
                    ' Start to edit the comment.
                    Dim commentDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = comment.BeginUpdate()
                    ' Insert a text to the comment.
                    commentDocument.Paragraphs.Insert(commentDocument.Range.Start)
                    commentDocument.InsertText(commentDocument.Range.Start, "some text")
                    ' Insert a table to the comment.
                    commentDocument.Tables.Create(commentDocument.Range.End, 5, 4)
                    ' Finalize to edit the comment.
                    comment.EndUpdate(commentDocument)
                End If
            End If
#End Region  ' #EditCommentContent
        End Sub
    End Class
End Namespace

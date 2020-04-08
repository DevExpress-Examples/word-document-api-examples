Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Public Class CommentsActions
		 Private Sub CreateComment(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CreateComment"
			Dim document As Document = wordProcessor.Document
			wordProcessor.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim docRange As DocumentRange = document.Paragraphs(2).Range
			Dim commentAuthor As String = "Johnson Alphonso D"
			document.Comments.Create(docRange, commentAuthor, Date.Now)
'			#End Region ' #CreateComment
		 End Sub

		 Private Sub CreateNestedComment(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CreateNestedComment"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			If document.Comments.Count > 0 Then
				Dim resRanges() As DocumentRange = document.FindAll("trump", SearchOptions.None, document.Comments(1).Range)
				If resRanges.Length > 0 Then
					Dim newComment As Comment = document.Comments.Create("Vicars Anny", document.Comments(1))
					newComment.Date = Date.Now
				End If
			End If
'			#End Region ' #CreateNestedComment
		 End Sub

		 Private Sub DeleteComment(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#DeleteComment"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			If document.Comments.Count > 0 Then
				document.Comments.Remove(document.Comments(0))
			End If
'			#End Region ' #DeleteComment
		 End Sub

		 Private Sub EditCommentProperties(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#EditCommentProperties"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim commentCount As Integer = document.Comments.Count
			If commentCount > 0 Then
				document.BeginUpdate()
				Dim comment As Comment = document.Comments(document.Comments.Count - 1)
				comment.Name = "New Name"
				comment.Date = Date.Now
				comment.Author = "New Author"
				document.EndUpdate()
			End If
'			#End Region ' #EditCommentProperties
		 End Sub

		 Private Sub EditCommentContent(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#EditCommentContent"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim commentCount As Integer = document.Comments.Count
			If commentCount > 0 Then
				Dim comment As Comment = document.Comments(document.Comments.Count - 1)
				If comment IsNot Nothing Then
					Dim commentDocument As SubDocument = comment.BeginUpdate()
					commentDocument.InsertText(commentDocument.CreatePosition(0), "some text")
					commentDocument.Tables.Create(commentDocument.CreatePosition(9), 5, 4)
					comment.EndUpdate(commentDocument)
				End If
			End If
'			#End Region ' #EditCommentContent
		 End Sub


	End Class
End Namespace

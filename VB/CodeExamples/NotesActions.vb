Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Public NotInheritable Class NotesActions

		Private Sub New()
		End Sub


		Private Shared Sub InsertFootnotes(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#InsertFootnotes"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Insert a footnote at the end of the 6th paragraph:
			Dim footnotePosition As DocumentPosition = document.CreatePosition(document.Paragraphs(5).Range.End.ToInt() - 1)
			document.Footnotes.Insert(footnotePosition)

			'Insert a footnote at the end of the 8th paragraph with a custom mark:
			Dim footnoteWithCustomMarkPosition As DocumentPosition = document.CreatePosition(document.Paragraphs(7).Range.End.ToInt() - 1)
			document.Footnotes.Insert(footnoteWithCustomMarkPosition, ChrW(&H00BA).ToString())
'			#End Region ' #InsertFootnotes 
		End Sub


		Private Shared Sub InsertEndnotes(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#InsertEndnotes"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Insert an endnote at the end of the last paragraph:
			Dim endnotePosition As DocumentPosition = document.CreatePosition(document.Paragraphs(document.Paragraphs.Count - 1).Range.End.ToInt() - 1)
			document.Endnotes.Insert(endnotePosition)

			'Insert an endnote at the end of the second last paragraph with a custom mark:
			Dim endnoteWithCustomMarkPosition As DocumentPosition = document.CreatePosition(document.Paragraphs(document.Paragraphs.Count - 2).Range.End.ToInt() - 1)
			document.Endnotes.Insert(endnoteWithCustomMarkPosition, ChrW(&H0060).ToString())
'			#End Region ' #InsertEndnotes
		End Sub

		Private Shared Sub EditFootnote(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#EditFootnote"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Access the first footnote's content:
			Dim footnote As SubDocument = document.Footnotes(0).BeginUpdate()

			'Exclude the reference mark and the space after it from the range to be edited:
			Dim noteTextRange As DocumentRange = footnote.CreateRange(footnote.Range.Start.ToInt() + 2, footnote.Range.Length - 2)

			'Clear the range:
			footnote.Delete(noteTextRange)

			'Append a new text:
			footnote.AppendText("the text is removed")

			'Finalize the update:
			document.Footnotes(0).EndUpdate(footnote)
'			#End Region ' #EditFootnote
		End Sub

		Private Shared Sub EditEndnote(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#EditEndnote"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Access the first endnote's content:
			Dim endnote As SubDocument = document.Endnotes(0).BeginUpdate()

			'Exclude the reference mark and the space after it from the range to be edited:
			Dim noteTextRange As DocumentRange = endnote.CreateRange(endnote.Range.Start.ToInt() + 2, endnote.Range.Length - 2)

			'Access the range's character properties:
			Dim characterProperties As CharacterProperties = endnote.BeginUpdateCharacters(noteTextRange)

			characterProperties.ForeColor = System.Drawing.Color.Red
			characterProperties.Italic = True

			'Finalize the character options update:
			endnote.EndUpdateCharacters(characterProperties)

			'Finalize the endnote update:
			document.Endnotes(0).EndUpdate(endnote)
'			#End Region ' #EditEndnote
		End Sub

		Private Shared Sub EditSeparator(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#EditSeparator"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Check whether the footnotes already have a separator:
			If document.Footnotes.HasSeparator(NoteSeparatorType.Separator) Then
				'Initiate the update session:
				Dim noteSeparator As SubDocument = document.Footnotes.BeginUpdateSeparator(NoteSeparatorType.Separator)

				'Clear the separator range:
				noteSeparator.Delete(noteSeparator.Range)

				'Append a new text:
				noteSeparator.AppendText("***")

				'Finalize the update:
				document.Footnotes.EndUpdateSeparator(noteSeparator)
			End If
'			#End Region ' #EditSeparator
		End Sub
		Private Shared Sub RemoveNotes(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#RemoveNotes"
			wordProcessor.LoadDocument("Documents//Grimm.docx")
			Dim document As Document = wordProcessor.Document

			'Remove first footnote:
			document.Footnotes.RemoveAt(0)


			'Remove all custom endnotes:
			For i As Integer = document.Endnotes.Count - 1 To 0 Step -1
				If document.Endnotes(i).IsCustom Then
					document.Endnotes.Remove(document.Endnotes(i))
				End If
			Next i

'			#End Region ' #RemoveNotes
		End Sub
	End Class
End Namespace

Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module NotesActions

        Private Sub InsertFootnotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertFootnotes"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Insert a footnote at the end of the 6th paragraph:
            Dim footnotePosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((5))).Range.[End].ToInt() - 1)
            document.Footnotes.Insert(footnotePosition)
            'Insert a footnote at the end of the 8th paragraph with a custom mark:
            Dim footnoteWithCustomMarkPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((7))).Range.[End].ToInt() - 1)
            document.Footnotes.Insert(footnoteWithCustomMarkPosition, "º")
#End Region  ' #InsertFootnotes 
        End Sub

        Private Sub InsertEndnotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertEndnotes"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Insert an endnote at the end of the last paragraph:
            Dim endnotePosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((document.Paragraphs.Count - 1))).Range.[End].ToInt() - 1)
            document.Endnotes.Insert(endnotePosition)
            'Insert an endnote at the end of the second last paragraph with a custom mark:
            Dim endnoteWithCustomMarkPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((document.Paragraphs.Count - 2))).Range.[End].ToInt() - 1)
            document.Endnotes.Insert(endnoteWithCustomMarkPosition, "`")
#End Region  ' #InsertEndnotes
        End Sub

        Private Sub EditFootnote(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditFootnote"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Access the first footnote's content:
            Dim footnote As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Footnotes(CInt((0))).BeginUpdate()
            'Exclude the reference mark and the space after it from the range to be edited:
            Dim noteTextRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = footnote.CreateRange(footnote.Range.Start.ToInt() + 2, footnote.Range.Length - 2)
            'Clear the range:
            footnote.Delete(noteTextRange)
            'Append a new text:
            footnote.AppendText("the text is removed")
            'Finalize the update:
            document.Footnotes(CInt((0))).EndUpdate(footnote)
#End Region  ' #EditFootnote
        End Sub

        Private Sub EditEndnote(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditEndnote"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Access the first endnote's content:
            Dim endnote As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Endnotes(CInt((0))).BeginUpdate()
            'Exclude the reference mark and the space after it from the range to be edited:
            Dim noteTextRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = endnote.CreateRange(endnote.Range.Start.ToInt() + 2, endnote.Range.Length - 2)
            'Access the range's character properties:
            Dim characterProperties As DevExpress.XtraRichEdit.API.Native.CharacterProperties = endnote.BeginUpdateCharacters(noteTextRange)
            characterProperties.ForeColor = System.Drawing.Color.Red
            characterProperties.Italic = True
            'Finalize the character options update:
            endnote.EndUpdateCharacters(characterProperties)
            'Finalize the endnote update:
            document.Endnotes(CInt((0))).EndUpdate(endnote)
#End Region  ' #EditEndnote
        End Sub

        Private Sub EditSeparator(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditSeparator"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Check whether the footnotes already have a separator:
            If document.Footnotes.HasSeparator(DevExpress.XtraRichEdit.API.Native.NoteSeparatorType.Separator) Then
                'Initiate the update session:
                Dim noteSeparator As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Footnotes.BeginUpdateSeparator(DevExpress.XtraRichEdit.API.Native.NoteSeparatorType.Separator)
                'Clear the separator range:
                noteSeparator.Delete(noteSeparator.Range)
                'Append a new text:
                noteSeparator.AppendText("***")
                'Finalize the update:
                document.Footnotes.EndUpdateSeparator(noteSeparator)
            End If
#End Region  ' #EditSeparator
        End Sub

        Private Sub RemoveNotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#RemoveNotes"
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            'Remove first footnote:
            document.Footnotes.RemoveAt(0)
            'Remove all custom endnotes:
            For i As Integer = document.Endnotes.Count - 1 To 0 Step -1
                If document.Endnotes(CInt((i))).IsCustom Then document.Endnotes.Remove(document.Endnotes(i))
            Next
#End Region  ' #RemoveNotes
        End Sub
    End Module
End Namespace

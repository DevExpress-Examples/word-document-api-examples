Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module NotesActions

        Public InsertFootnotesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.InsertFootnotes

        Public InsertEndnotesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.InsertEndnotes

        Public EditFootnoteAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.EditFootnote

        Public EditEndnoteAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.EditEndnote

        Public EditSeparatorAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.EditSeparator

        Public RemoveNotesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.NotesActions.RemoveNotes

        Private Sub InsertFootnotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertFootnotes"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Paragraphs.Count > 5 Then
                ' Insert a footnote at the end of the sixth paragraph.
                Dim footnotePosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((5))).Range.[End].ToInt() - 1)
                document.Footnotes.Insert(footnotePosition)
                ' Insert a footnote at the end of the eighth paragraph with a custom mark.
                Dim footnoteWithCustomMarkPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((7))).Range.[End].ToInt() - 1)
                document.Footnotes.Insert(footnoteWithCustomMarkPosition, "º")
            End If
#End Region  ' #InsertFootnotes 
        End Sub

        Private Sub InsertEndnotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertEndnotes"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Insert an endnote at the end of the last paragraph.
            Dim endnotePosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((document.Paragraphs.Count - 1))).Range.[End].ToInt() - 1)
            document.Endnotes.Insert(endnotePosition)
            ' Insert an endnote at the end of the second last paragraph with a custom mark.
            Dim endnoteWithCustomMarkPosition As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition(document.Paragraphs(CInt((document.Paragraphs.Count - 2))).Range.[End].ToInt() - 1)
            document.Endnotes.Insert(endnoteWithCustomMarkPosition, "`")
#End Region  ' #InsertEndnotes
        End Sub

        Private Sub EditFootnote(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditFootnote"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the first footnote content.
            Dim footnote As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Footnotes(CInt((0))).BeginUpdate()
            ' Exclude the reference mark and the space after it from the range that is edited.
            Dim noteTextRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = footnote.CreateRange(footnote.Range.Start.ToInt() + 2, footnote.Range.Length - 2)
            ' Clear the range.
            footnote.Delete(noteTextRange)
            ' Change the footnote text.
            footnote.AppendText("the text is removed")
            ' Finalize to update the endnote.
            document.Footnotes(CInt((0))).EndUpdate(footnote)
#End Region  ' #EditFootnote
        End Sub

        Private Sub EditEndnote(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditEndnote"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the first endnote content.
            Dim endnote As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Endnotes(CInt((0))).BeginUpdate()
            ' Exclude the reference mark and the space after it from the range that is edited.
            Dim noteTextRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = endnote.CreateRange(endnote.Range.Start.ToInt() + 2, endnote.Range.Length - 2)
            ' Access the endnote's character formatting.
            Dim characterProperties As DevExpress.XtraRichEdit.API.Native.CharacterProperties = endnote.BeginUpdateCharacters(noteTextRange)
            ' Specify the endnote's character formatting options.
            characterProperties.ForeColor = System.Drawing.Color.Red
            characterProperties.Italic = True
            ' Finalize to update character formatting.
            endnote.EndUpdateCharacters(characterProperties)
            ' Finalize to update the endnote.
            document.Endnotes(CInt((0))).EndUpdate(endnote)
#End Region  ' #EditEndnote
        End Sub

        Private Sub EditSeparator(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#EditSeparator"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check whether footnotes already have a separator.
            If document.Footnotes.HasSeparator(DevExpress.XtraRichEdit.API.Native.NoteSeparatorType.Separator) Then
                ' Access the footnote separator.
                Dim noteSeparator As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Footnotes.BeginUpdateSeparator(DevExpress.XtraRichEdit.API.Native.NoteSeparatorType.Separator)
                ' Clear the separator range.
                noteSeparator.Delete(noteSeparator.Range)
                ' Change the footnote separator.
                noteSeparator.AppendText("***")
                ' Finalize to update the footnote separator.
                document.Footnotes.EndUpdateSeparator(noteSeparator)
            End If
#End Region  ' #EditSeparator
        End Sub

        Private Sub RemoveNotes(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#RemoveNotes"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx")
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Remove the first footnote.
            If document.Footnotes.Count > 0 Then document.Footnotes.RemoveAt(0)
            ' Remove all custom endnotes.
            For i As Integer = document.Endnotes.Count - 1 To 0 Step -1
                If document.Endnotes(CInt((i))).IsCustom Then document.Endnotes.Remove(document.Endnotes(i))
            Next
#End Region  ' #RemoveNotes
        End Sub
    End Module
End Namespace

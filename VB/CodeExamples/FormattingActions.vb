Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Drawing

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FormattingActions

        Public Shared FormatTextAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormattingActions.FormatText

        Public Shared ChangeSpacingAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormattingActions.ChangeSpacing

        Public Shared ResetCharacterFormattingAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormattingActions.ResetCharacterFormatting

        Public Shared FormatParagraphAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormattingActions.FormatParagraph

        Public Shared ResetParagraphFormattingAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.FormattingActions.ResetParagraphFormatting

        Private Shared Sub FormatText(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#FormatText"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText("Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Formatted" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Access the range of the document's second paragraph.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((1))).Range
            ' Start to modify character formatting of the target range.
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            ' Specify character formatting options.
            cp.FontName = "Comic Sans MS"
            cp.FontSize = 18
            cp.ForeColor = System.Drawing.Color.Blue
            cp.BackColor = System.Drawing.Color.Snow
            cp.Underline = DevExpress.XtraRichEdit.API.Native.UnderlineType.DoubleWave
            cp.UnderlineColor = System.Drawing.Color.Red
            ' Finalize to modify character formatting.
            document.EndUpdateCharacters(cp)
#End Region  ' #FormatText
        End Sub

        Private Shared Sub ChangeSpacing(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeCharacterSpacing"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText("Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Formatted" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Access the range of the document's second paragraph.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((1))).Range
            ' Start to modify character formatting of the target range.
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            ' Change character spacing and scaling.
            cp.Scale = 150
            cp.Spacing = -2
            ' Raise the text by 2 points.
            cp.Position = 2
            ' Finalize to modify character formatting.
            document.EndUpdateCharacters(cp)
#End Region  ' #ChangeCharacterSpacing
        End Sub

        Private Shared Sub ResetCharacterFormatting(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ResetCharacterFormatting"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the range of the document's first paragraph.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            ' Start to modify character formatting of the target range.
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            ' Set the font size and font name of the target range's characters to default values.   
            ' Other character properties remain intact.
            cp.Reset(DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontSize Or DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontName Or DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontNameAscii)
            ' Finalize to modify character formatting.
            document.EndUpdateCharacters(cp)
#End Region  ' #ResetCharacterFormatting
        End Sub

        Private Shared Sub FormatParagraph(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#FormatParagraph"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText("Modified Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Access the first paragraph range.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            ' Start to edit the paragraph.
            Dim pp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Specify the paragraph's alignment.
            pp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
            ' Specify the paragraph's line spacing.
            pp.LineSpacingType = DevExpress.XtraRichEdit.API.Native.ParagraphLineSpacing.Multiple
            pp.LineSpacingMultiplier = 3
            ' Set the paragraph’s left indent to 0.5 document unit.
            ' Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5F)
            ' Start to modify tab stops in the paragraph.
            Dim tbiColl As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = pp.BeginUpdateTabs(True)
            ' Create a new tab stop for the paragraph.
            Dim tbi As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            ' Specify the tab stop's alignment type.
            tbi.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.Center
            ' Set the tab stop position to 1.5 document unit.
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            ' Add the tab stop to the collection of tab stops.
            tbiColl.Add(tbi)
            ' Finalize to modify tab stops in the paragraph.
            pp.EndUpdateTabs(tbiColl)
            ' Finalize to edit the paragraph.
            document.EndUpdateParagraphs(pp)
#End Region  ' #FormatParagraph
        End Sub

        Private Shared Sub ResetParagraphFormatting(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ResetParagraphFormatting"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the range of the document's first paragraph.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            ' Start to edit the paragraph.
            Dim cp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Set alignmment and first line indent of the target paragraph to default values.   
            ' Other paragraph properties remain intact.
            cp.Reset(DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.Alignment Or DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.FirstLineIndent)
            ' Finalize to edit the paragraph.
            document.EndUpdateParagraphs(cp)
#End Region  ' #ResetParagraphFormatting
        End Sub
    End Class
End Namespace

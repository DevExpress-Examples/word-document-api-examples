Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class FormattingActions

        Private Shared Sub FormatText(ByVal server As RichEditDocumentServer)
'            #Region "#FormatText"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.AppendText("Normal" & ControlChars.Lf & "Formatted" & ControlChars.Lf & "Normal")
            document.EndUpdate()
            Dim range As DocumentRange = document.Paragraphs(1).Range
            Dim cp As CharacterProperties = document.BeginUpdateCharacters(range)
            cp.FontName = "Comic Sans MS"
            cp.FontSize = 18
            cp.ForeColor = Color.Blue
            cp.BackColor = Color.Snow
            cp.Underline = UnderlineType.DoubleWave
            cp.UnderlineColor = Color.Red
            document.EndUpdateCharacters(cp)
'            #End Region ' #FormatText
        End Sub
        Private Shared Sub ResetCharacterFormatting(ByVal server As RichEditDocumentServer)
'            #Region "#ResetCharacterFormatting"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            ' Set font size and font name of the characters in the first paragraph to default. 
            ' Other character properties remain intact.
            Dim range As DocumentRange = document.Paragraphs(0).Range
            Dim cp As CharacterProperties = document.BeginUpdateCharacters(range)
            cp.Reset(CharacterPropertiesMask.FontSize Or CharacterPropertiesMask.FontName)
            document.EndUpdateCharacters(cp)
'            #End Region ' #ResetCharacterFormatting
        End Sub
        Private Shared Sub FormatParagraph(ByVal server As RichEditDocumentServer)
'            #Region "#FormatParagraph"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.AppendText("Modified Paragraph" & ControlChars.Lf & "Normal" & ControlChars.Lf & "Normal")
            document.EndUpdate()
            Dim pos As DocumentPosition = document.Range.Start
            Dim range As DocumentRange = document.CreateRange(pos, 0)
            Dim pp As ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Center paragraph
            pp.Alignment = ParagraphAlignment.Center
            ' Set triple spacing
            pp.LineSpacingType = ParagraphLineSpacing.Multiple
            pp.LineSpacingMultiplier = 3
            ' Set left indent at 0.5".
            ' Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5F)
            ' Set tab stop at 1.5"
            Dim tbiColl As TabInfoCollection = pp.BeginUpdateTabs(True)
            Dim tbi As TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            tbi.Alignment = TabAlignmentType.Center
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            tbiColl.Add(tbi)
            pp.EndUpdateTabs(tbiColl)
            document.EndUpdateParagraphs(pp)
'            #End Region ' #FormatParagraph
        End Sub
        Private Shared Sub ResetParagraphFormatting(ByVal server As RichEditDocumentServer)
'            #Region "#ResetParagraphFormatting"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            ' Set alignment and indentation of the first line in the first paragraph to default. 
            ' Other paragraph properties remain intact.
            Dim range As DocumentRange = document.Paragraphs(0).Range
            Dim cp As ParagraphProperties = document.BeginUpdateParagraphs(range)
            cp.Reset(ParagraphPropertiesMask.Alignment Or ParagraphPropertiesMask.FirstLineIndent)
            document.EndUpdateParagraphs(cp)
'            #End Region ' #ResetParagraphFormatting
        End Sub
    End Class
End Namespace

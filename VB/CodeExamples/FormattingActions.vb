Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class FormattingActions

        Private Shared Sub FormatText(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#FormatText"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.AppendText("Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Formatted" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            document.EndUpdate()
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((1))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            cp.FontName = "Comic Sans MS"
            cp.FontSize = 18
            cp.ForeColor = System.Drawing.Color.Blue
            cp.BackColor = System.Drawing.Color.Snow
            cp.Underline = DevExpress.XtraRichEdit.API.Native.UnderlineType.DoubleWave
            cp.UnderlineColor = System.Drawing.Color.Red
            document.EndUpdateCharacters(cp)
#End Region  ' #FormatText
        End Sub

        Private Shared Sub ChangeSpacing(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeCharacterSpacing"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.AppendText("Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Formatted" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            document.EndUpdate()
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            cp.Scale = 150
            cp.Spacing = -2
            cp.Position = 2
            document.EndUpdateCharacters(cp)
#End Region  ' #ChangeCharacterSpacing
        End Sub

        Private Shared Sub ResetCharacterFormatting(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ResetCharacterFormatting"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Set font size and font name of the characters in the first paragraph to default. 
            ' Other character properties remain intact.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(range)
            cp.Reset(DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontSize Or DevExpress.XtraRichEdit.API.Native.CharacterPropertiesMask.FontName)
            document.EndUpdateCharacters(cp)
#End Region  ' #ResetCharacterFormatting
        End Sub

        Private Shared Sub FormatParagraph(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#FormatParagraph"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            document.AppendText("Modified Paragraph" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal" & Global.Microsoft.VisualBasic.Constants.vbLf & "Normal")
            document.EndUpdate()
            Dim pos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Range.Start
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(pos, 0)
            Dim pp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            ' Center paragraph
            pp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
            ' Set triple spacing
            pp.LineSpacingType = DevExpress.XtraRichEdit.API.Native.ParagraphLineSpacing.Multiple
            pp.LineSpacingMultiplier = 3
            ' Set left indent at 0.5".
            ' Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5F)
            ' Set tab stop at 1.5"
            Dim tbiColl As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = pp.BeginUpdateTabs(True)
            Dim tbi As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            tbi.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.Center
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            tbiColl.Add(tbi)
            pp.EndUpdateTabs(tbiColl)
            document.EndUpdateParagraphs(pp)
#End Region  ' #FormatParagraph
        End Sub

        Private Shared Sub ResetParagraphFormatting(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ResetParagraphFormatting"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Set alignment and indentation of the first line in the first paragraph to default. 
            ' Other paragraph properties remain intact.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            Dim cp As DevExpress.XtraRichEdit.API.Native.ParagraphProperties = document.BeginUpdateParagraphs(range)
            cp.Reset(DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.Alignment Or DevExpress.XtraRichEdit.API.Native.ParagraphPropertiesMask.FirstLineIndent)
            document.EndUpdateParagraphs(cp)
#End Region  ' #ResetParagraphFormatting
        End Sub
    End Class
End Namespace

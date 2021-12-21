Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class StylesAction

        Public Shared CreateNewCharacterStyleAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.StylesAction.CreateNewCharacterStyle

        Public Shared CreateNewParagraphStyleAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.StylesAction.CreateNewParagraphStyle

        Public Shared CreateNewLinkedStyleAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.StylesAction.CreateNewLinkedStyle

        Private Shared Sub CreateNewCharacterStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNewCharacterStyle"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the character style with the specified name.
            Dim cstyle As DevExpress.XtraRichEdit.API.Native.CharacterStyle = document.CharacterStyles("MyCStyle")
            ' If the style with the specified name does not exist
            ' create a new character style and specify the style settings.
            If cstyle Is Nothing Then
                cstyle = document.CharacterStyles.CreateNew()
                cstyle.Name = "MyCStyle"
                cstyle.Parent = document.CharacterStyles("Default Paragraph Font")
                cstyle.ForeColor = System.Drawing.Color.DarkOrange
                cstyle.Strikeout = DevExpress.XtraRichEdit.API.Native.StrikeoutType.[Double]
                cstyle.FontName = "Verdana"
                ' Add the style to the collection of character styles.
                document.CharacterStyles.Add(cstyle)
            End If

            ' Access the range of the first paragraph.
            Dim myRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.Paragraphs(CInt((0))).Range
            ' Access character formatting of the target range.
            Dim charProps As DevExpress.XtraRichEdit.API.Native.CharacterProperties = document.BeginUpdateCharacters(myRange)
            ' Apply the created character style to the target range.
            charProps.Style = cstyle
            ' Finalize to modify character formatting.
            document.EndUpdateCharacters(charProps)
#End Region  ' #CreateNewCharacterStyle
        End Sub

        Private Shared Sub CreateNewParagraphStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNewParagraphStyle"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access a paragraph style with the specified name.
            Dim pstyle As DevExpress.XtraRichEdit.API.Native.ParagraphStyle = document.ParagraphStyles("MyPStyle")
            ' If the style with the specified name does not exist
            ' create a new paragraph style and specify the style settings.
            If pstyle Is Nothing Then
                pstyle = document.ParagraphStyles.CreateNew()
                pstyle.Name = "MyPStyle"
                pstyle.LineSpacingType = DevExpress.XtraRichEdit.API.Native.ParagraphLineSpacing.[Double]
                pstyle.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
                ' Add the style to the collection of paragraph styles.
                document.ParagraphStyles.Add(pstyle)
            End If

            If document.Paragraphs.Count > 2 Then
                ' Apply the created paragraph style to the third document paragraph.
                document.Paragraphs(CInt((2))).Style = pstyle
            End If
#End Region  ' #CreateNewParagraphStyle
        End Sub

        Private Shared Sub CreateNewLinkedStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNewLinkedStyle"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Append text to the document.
            document.AppendText("Line One" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Two" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Three")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Access a paragraph style with the specified name.
            Dim lstyle As DevExpress.XtraRichEdit.API.Native.ParagraphStyle = document.ParagraphStyles("MyLinkedStyle")
            ' If the style with the specified name does not exist
            ' create a new paragraph and character styles and specify their settings.
            If lstyle Is Nothing Then
                ' Start to edit the document.
                document.BeginUpdate()
                ' Create a paragraph style and specify its settings.
                lstyle = document.ParagraphStyles.CreateNew()
                lstyle.Name = "MyLinkedStyle"
                lstyle.LineSpacingType = DevExpress.XtraRichEdit.API.Native.ParagraphLineSpacing.[Double]
                lstyle.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
                document.ParagraphStyles.Add(lstyle)
                ' Create a character style and specify its settings.
                Dim lcstyle As DevExpress.XtraRichEdit.API.Native.CharacterStyle = document.CharacterStyles.CreateNew()
                lcstyle.Name = "MyLinkedCStyle"
                document.CharacterStyles.Add(lcstyle)
                ' Set the created character style to the created paragraph style.
                lcstyle.LinkedStyle = lstyle
                ' Specify the created character style's settings.
                lcstyle.ForeColor = System.Drawing.Color.DarkGreen
                lcstyle.Strikeout = DevExpress.XtraRichEdit.API.Native.StrikeoutType.[Single]
                lcstyle.FontSize = 24
                ' Finalize to edit the document.
                document.EndUpdate()
                ' Save the resulting document and select it in the File Explorer.
                document.SaveDocument("LinkedStyleSample.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start("explorer.exe", "/select," & "LinkedStyleSample.docx")
            End If
#End Region  ' #CreateNewLinkedStyle
        End Sub
    End Class
End Namespace

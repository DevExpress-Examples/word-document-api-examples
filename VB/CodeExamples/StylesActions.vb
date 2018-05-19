Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RTEDocumentServerExamples.CodeExamples
    Friend Class StylesAction
        Private Shared Sub CreateNewCharacterStyle(ByVal server As RichEditDocumentServer)
'            #Region "#CreateNewCharacterStyle"
            Dim document As Document = server.Document
            server.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            Dim cstyle As CharacterStyle = document.CharacterStyles("MyCStyle")
            If cstyle Is Nothing Then
                cstyle = document.CharacterStyles.CreateNew()
                cstyle.Name = "MyCStyle"
                cstyle.Parent = document.CharacterStyles("Default Paragraph Font")
                cstyle.ForeColor = System.Drawing.Color.DarkOrange
                cstyle.Strikeout = StrikeoutType.Double
                cstyle.FontName = "Verdana"
                document.CharacterStyles.Add(cstyle)
            End If
            Dim myRange As DocumentRange = document.Paragraphs(0).Range
            Dim charProps As CharacterProperties = document.BeginUpdateCharacters(myRange)
            charProps.Style = cstyle
            document.EndUpdateCharacters(charProps)
'            #End Region ' #CreateNewCharacterStyle
        End Sub

        Private Shared Sub CreateNewParagraphStyle(ByVal server As RichEditDocumentServer)
'            #Region "#CreateNewParagraphStyle"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim pstyle As ParagraphStyle = document.ParagraphStyles("MyPStyle")
            If pstyle Is Nothing Then
                pstyle = document.ParagraphStyles.CreateNew()
                pstyle.Name = "MyPStyle"
                pstyle.LineSpacingType = ParagraphLineSpacing.Double
                pstyle.Alignment = ParagraphAlignment.Center
                document.ParagraphStyles.Add(pstyle)
            End If
            document.Paragraphs(2).Style = pstyle
'            #End Region ' #CreateNewParagraphStyle
        End Sub

        Private Shared Sub CreateNewLinkedStyle(ByVal server As RichEditDocumentServer)
'            #Region "#CreateNewLinkedStyle"
            Dim document As Document = server.Document
            document.BeginUpdate()
            document.AppendText("Line One" & ControlChars.Lf & "Line Two" & ControlChars.Lf & "Line Three")
            document.EndUpdate()
            Dim lstyle As ParagraphStyle = document.ParagraphStyles("MyLinkedStyle")
            If lstyle Is Nothing Then
                document.BeginUpdate()
                lstyle = document.ParagraphStyles.CreateNew()
                lstyle.Name = "MyLinkedStyle"
                lstyle.LineSpacingType = ParagraphLineSpacing.Double
                lstyle.Alignment = ParagraphAlignment.Center
                document.ParagraphStyles.Add(lstyle)

                Dim lcstyle As CharacterStyle = document.CharacterStyles.CreateNew()
                lcstyle.Name = "MyLinkedCStyle"
                document.CharacterStyles.Add(lcstyle)
                lcstyle.LinkedStyle = lstyle

                lcstyle.ForeColor = System.Drawing.Color.DarkGreen
                lcstyle.Strikeout = StrikeoutType.Single
                lcstyle.FontSize = 24
                document.EndUpdate()
                document.SaveDocument("LinkedStyleSample.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start("explorer.exe", "/select," & "LinkedStyleSample.docx")
            End If
'            #End Region ' #CreateNewLinkedStyle
        End Sub
    End Class
End Namespace

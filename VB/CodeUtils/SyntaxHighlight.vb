Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports DevExpress.CodeParser
Imports DevExpress.Office.Internal
Imports DevExpress.Office.Utils
Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Import
Imports DevExpress.XtraRichEdit.Internal
Imports DevExpress.XtraRichEdit.Services

Namespace RichEditDocumentServerAPIExample.CodeUtils

    Public Class SyntaxHightlightInitializeHelper

        Public Sub Initialize(ByVal richEditControl As IRichEditControl, ByVal codeExamplesFileExtension As String)
            Dim innerControl As InnerRichEditControl = richEditControl.InnerControl
            Dim commandFactory As IRichEditCommandFactoryService = innerControl.GetService(Of IRichEditCommandFactoryService)()
            If commandFactory Is Nothing Then Return ' wpf richedit is not loaded
            innerControl.ReplaceService(Of ISyntaxHighlightService)(New SyntaxHighlightService(innerControl, codeExamplesFileExtension))
            Dim newCommandFactory As CustomRichEditCommandFactoryService = New CustomRichEditCommandFactoryService(commandFactory)
            innerControl.RemoveService(GetType(IRichEditCommandFactoryService))
            innerControl.AddService(GetType(IRichEditCommandFactoryService), newCommandFactory)
            Dim importManager As IDocumentImportManagerService = innerControl.GetService(Of IDocumentImportManagerService)()
            importManager.UnregisterAllImporters()
            importManager.RegisterImporter(New PlainTextDocumentImporter())
            importManager.RegisterImporter(New SourcesCodeDocumentImporter())
            Dim exportManager As IDocumentExportManagerService = innerControl.GetService(Of IDocumentExportManagerService)()
            exportManager.UnregisterAllExporters()
            exportManager.RegisterExporter(New PlainTextDocumentExporter())
            exportManager.RegisterExporter(New SourcesCodeDocumentExporter())
            Dim document As Document = innerControl.Document
            document.BeginUpdate()
            Try
                document.DefaultCharacterProperties.FontName = "Consolas"
                document.DefaultCharacterProperties.FontSize = 10
                document.Sections(0).Page.Width = Units.InchesToDocumentsF(20)
                document.Sections(0).Margins.Top = Units.InchesToDocumentsF(0.2F)
                document.Sections(0).Margins.Left = Units.InchesToDocumentsF(0.2F)
                document.Sections(0).Margins.Right = Units.InchesToDocumentsF(0.2F)
            Finally
                document.EndUpdate()
            End Try
        End Sub
    End Class

    Public Class SyntaxHighlightService
        Implements ISyntaxHighlightService

        Private ReadOnly editor As InnerRichEditControl

        Private ReadOnly syntaxHighlightInfo As SyntaxHighlightInfo

        Private ReadOnly fileExtensionToHightlight As String

        Public Sub New(ByVal editor As InnerRichEditControl, ByVal extension As String)
            Me.editor = editor
            syntaxHighlightInfo = New SyntaxHighlightInfo()
            fileExtensionToHightlight = extension
        End Sub

        Private Sub ForceExecute() Implements ISyntaxHighlightService.ForceExecute
            ExecuteCore()
        End Sub

        Private Sub Execute() Implements ISyntaxHighlightService.Execute
            ExecuteCore()
        End Sub

        Private Sub ExecuteCore()
            Dim tokens As TokenCollection = Parse(editor.Text)
            HighlightSyntax(tokens)
        End Sub

        Private Function Parse(ByVal code As String) As TokenCollection
            If String.IsNullOrEmpty(code) Then
                Return Nothing
            End If

            Dim tokenizer As ITokenCategoryHelper = CreateTokenizer()
            If tokenizer Is Nothing Then
                Return New TokenCollection()
            End If

            Return tokenizer.GetTokens(code)
        End Function

        Private Function CreateTokenizer() As ITokenCategoryHelper
            Dim fileName As String = editor.Options.DocumentSaveOptions.CurrentFileName
            Dim extenstion As String
            If String.IsNullOrEmpty(fileName) Then
                extenstion = fileExtensionToHightlight
            Else
                extenstion = Path.GetExtension(fileName)
            End If

            Dim result As ITokenCategoryHelper = TokenCategoryHelperFactory.CreateHelperForFileExtensions(extenstion)
            If result IsNot Nothing Then
                Return result
            Else
                Return Nothing
            End If
        End Function

        Private Sub HighlightSyntax(ByVal tokens As TokenCollection)
            If tokens Is Nothing OrElse tokens.Count = 0 Then
                Return
            End If

            Dim document As Document = editor.Document
            Dim cp As CharacterProperties = document.BeginUpdateCharacters(0, 1)
            Dim syntaxTokens As List(Of SyntaxHighlightToken) = New List(Of SyntaxHighlightToken)(tokens.Count)
            For Each token As Token In tokens
                HighlightCategorizedToken(CType(token, CategorizedToken), syntaxTokens)
            Next

            document.ApplySyntaxHighlight(syntaxTokens)
            document.EndUpdateCharacters(cp)
        End Sub

        Private Sub HighlightCategorizedToken(ByVal token As CategorizedToken, ByVal syntaxTokens As List(Of SyntaxHighlightToken))
            Dim backColor As Color = editor.ActiveView.BackColor
            Dim highlightProperties As SyntaxHighlightProperties = syntaxHighlightInfo.CalculateTokenCategoryHighlight(token.Category)
            Dim syntaxToken As SyntaxHighlightToken = SetTokenColor(token, highlightProperties, backColor)
            If syntaxToken IsNot Nothing Then
                syntaxTokens.Add(syntaxToken)
            End If
        End Sub

        Private Function SetTokenColor(ByVal token As Token, ByVal foreColor As SyntaxHighlightProperties, ByVal backColor As Color) As SyntaxHighlightToken
            If editor.Document.Paragraphs.Count < token.Range.Start.Line Then
                Return Nothing
            End If

            Dim paragraphStart As Integer = DocumentHelper.GetParagraphStart(editor.Document.Paragraphs(token.Range.Start.Line - 1))
            Dim tokenStart As Integer = paragraphStart + token.Range.Start.Offset - 1
            If token.Range.End.Line <> token.Range.Start.Line Then
                paragraphStart = DocumentHelper.GetParagraphStart(editor.Document.Paragraphs(token.Range.End.Line - 1))
            End If

            Dim tokenEnd As Integer = paragraphStart + token.Range.End.Offset - 1
            System.Diagnostics.Debug.Assert(tokenEnd > tokenStart)
            Return New SyntaxHighlightToken(tokenStart, tokenEnd - tokenStart, foreColor)
        End Function
    End Class

    Public Class SyntaxHighlightInfo

        Private ReadOnly properties As Dictionary(Of TokenCategory, SyntaxHighlightProperties)

        Public Sub New()
            properties = New Dictionary(Of TokenCategory, SyntaxHighlightProperties)()
            Reset()
        End Sub

        Public Sub Reset()
            properties.Clear()
            Add(TokenCategory.Text, DXColor.Black)
            Add(TokenCategory.Keyword, DXColor.Blue)
            Add(TokenCategory.String, DXColor.Brown)
            Add(TokenCategory.Comment, DXColor.Green)
            Add(TokenCategory.Identifier, DXColor.Black)
            Add(TokenCategory.PreprocessorKeyword, DXColor.Blue)
            Add(TokenCategory.Number, DXColor.Red)
            Add(TokenCategory.Operator, DXColor.Black)
            Add(TokenCategory.Unknown, DXColor.Black)
            Add(TokenCategory.XmlComment, DXColor.Gray)
            Add(TokenCategory.CssComment, DXColor.Green)
            Add(TokenCategory.CssKeyword, DXColor.Brown)
            Add(TokenCategory.CssPropertyName, DXColor.Red)
            Add(TokenCategory.CssPropertyValue, DXColor.Blue)
            Add(TokenCategory.CssSelector, DXColor.Blue)
            Add(TokenCategory.CssStringValue, DXColor.Blue)
            Add(TokenCategory.HtmlAttributeName, DXColor.Red)
            Add(TokenCategory.HtmlAttributeValue, DXColor.Blue)
            Add(TokenCategory.HtmlComment, DXColor.Green)
            Add(TokenCategory.HtmlElementName, DXColor.Brown)
            Add(TokenCategory.HtmlEntity, DXColor.Gray)
            Add(TokenCategory.HtmlOperator, DXColor.Black)
            Add(TokenCategory.HtmlServerSideScript, DXColor.Black)
            Add(TokenCategory.HtmlString, DXColor.Blue)
            Add(TokenCategory.HtmlTagDelimiter, DXColor.Blue)
        End Sub

        Private Sub Add(ByVal category As TokenCategory, ByVal foreColor As Color)
            Dim item As SyntaxHighlightProperties = New SyntaxHighlightProperties()
            item.ForeColor = foreColor
            properties.Add(category, item)
        End Sub

        Public Function CalculateTokenCategoryHighlight(ByVal category As TokenCategory) As SyntaxHighlightProperties
            Dim result As SyntaxHighlightProperties = CType(Nothing, SyntaxHighlightProperties)
            If properties.TryGetValue(category, result) Then
                Return result
            Else
                Return properties(TokenCategory.Text)
            End If
        End Function
    End Class

    Public Class CustomRichEditCommandFactoryService
        Implements IRichEditCommandFactoryService

        Private ReadOnly service As IRichEditCommandFactoryService

        Public Sub New(ByVal service As IRichEditCommandFactoryService)
            Guard.ArgumentNotNull(service, "service")
            Me.service = service
        End Sub

        Private Function CreateCommand(ByVal id As RichEditCommandId) As RichEditCommand Implements IRichEditCommandFactoryService.CreateCommand
            If id.Equals(RichEditCommandId.InsertColumnBreak) OrElse id.Equals(RichEditCommandId.InsertLineBreak) OrElse id.Equals(RichEditCommandId.InsertPageBreak) Then
                Return service.CreateCommand(RichEditCommandId.InsertParagraph)
            End If

            Return service.CreateCommand(id)
        End Function
    End Class

    Public Module SourceCodeDocumentFormat

        Public ReadOnly Id As DocumentFormat = New DocumentFormat(1325)
    End Module

    Public Class SourcesCodeDocumentImporter
        Inherits PlainTextDocumentImporter

        Friend Shared ReadOnly filterField As FileDialogFilter = New FileDialogFilter("Source Files", New String() {"cs", "vb", "html", "htm", "js", "xml", "css"})

        Public Overrides ReadOnly Property Filter As FileDialogFilter
            Get
                Return filterField
            End Get
        End Property

        Public Overrides ReadOnly Property Format As DocumentFormat
            Get
                Return Id
            End Get
        End Property
    End Class

    Public Class SourcesCodeDocumentExporter
        Inherits PlainTextDocumentExporter

        Public Overrides ReadOnly Property Filter As FileDialogFilter
            Get
                Return SourcesCodeDocumentImporter.filterField
            End Get
        End Property

        Public Overrides ReadOnly Property Format As DocumentFormat
            Get
                Return Id
            End Get
        End Property
    End Class
End Namespace

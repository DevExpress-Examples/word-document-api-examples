Imports System
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Internal

Namespace RichEditDocumentServerAPIExample.CodeUtils

    Public Class ExampleCodeEditor

        Private ReadOnly codeEditorCs As IRichEditControl

        Private ReadOnly codeEditorVb As IRichEditControl

        Private current As ExampleLanguage

        Public Sub New(ByVal codeEditorCs As IRichEditControl, ByVal codeEditorVb As IRichEditControl)
            Me.codeEditorCs = codeEditorCs
            Me.codeEditorVb = codeEditorVb
            AddHandler Me.codeEditorCs.InnerControl.InitializeDocument, New EventHandler(AddressOf InitializeSyntaxHighlightForCs)
            AddHandler Me.codeEditorVb.InnerControl.InitializeDocument, New EventHandler(AddressOf InitializeSyntaxHighlightForVb)
        End Sub

        Private Sub InitializeSyntaxHighlightForCs(ByVal sender As Object, ByVal e As EventArgs)
            InitializeSyntaxHighlight(codeEditorCs, ExampleLanguage.Csharp)
        End Sub

        Private Sub InitializeSyntaxHighlightForVb(ByVal sender As Object, ByVal e As EventArgs)
            InitializeSyntaxHighlight(codeEditorVb, ExampleLanguage.VB)
        End Sub

        Private Sub InitializeSyntaxHighlight(ByVal codeEditor As IRichEditControl, ByVal language As ExampleLanguage)
            Dim syntaxHightlightInitializator As SyntaxHightlightInitializeHelper = New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditor, GetCodeExampleFileExtension(language))
            DisableRichEditFeatures(codeEditor)
        End Sub

        Public ReadOnly Property CurrentCodeEditor As InnerRichEditControl
            Get
                If CurrentExampleLanguage = ExampleLanguage.Csharp Then
                    Return codeEditorCs.InnerControl
                Else
                    Return codeEditorVb.InnerControl
                End If
            End Get
        End Property

        Public Property CurrentExampleLanguage As ExampleLanguage
            Get
                Return current
            End Get

            Set(ByVal value As ExampleLanguage)
                current = value
            End Set
        End Property

        Public Sub ShowExample(ByVal codeExample As RichEditExample)
            Dim richEditControlCs As InnerRichEditControl = codeEditorCs.InnerControl
            Dim richEditControlVb As InnerRichEditControl = codeEditorVb.InnerControl
            If codeExample IsNot Nothing Then
                richEditControlCs.Text = codeExample.CodeCS
                richEditControlVb.Text = codeExample.CodeVB
            End If
        End Sub

        Private Sub DisableRichEditFeatures(ByVal codeEditor As IRichEditControl)
            Dim options As RichEditControlOptionsBase = codeEditor.InnerDocumentServer.Options
            options.DocumentCapabilities.Hyperlinks = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.Bulleted = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.Simple = DocumentCapability.Disabled
            options.DocumentCapabilities.Numbering.MultiLevel = DocumentCapability.Disabled
            options.DocumentCapabilities.Tables = DocumentCapability.Disabled
            options.DocumentCapabilities.Bookmarks = DocumentCapability.Disabled
            options.DocumentCapabilities.CharacterStyle = DocumentCapability.Disabled
            options.DocumentCapabilities.ParagraphStyle = DocumentCapability.Disabled
        End Sub
    End Class
End Namespace

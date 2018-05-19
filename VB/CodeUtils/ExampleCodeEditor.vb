Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Internal
Imports System

Namespace RichEditDocumentServerAPIExample.CodeUtils
    Public Class ExampleCodeEditor
        Private ReadOnly codeEditorCs As IRichEditControl
        Private ReadOnly codeEditorVb As IRichEditControl
        Private ReadOnly codeEditorCsClass As IRichEditControl
        Private ReadOnly codeEditorVbClass As IRichEditControl

        Private current As ExampleLanguage

        Private forceTextChangesCounter As Integer

        Private richEditTextChanged_Renamed As Boolean = False

        Private lastExampleCodeModifiedTime_Renamed As Date = Date.Now

        Public Sub New(ByVal codeEditorCs As IRichEditControl, ByVal codeEditorVb As IRichEditControl, ByVal codeEditorCsClass As IRichEditControl, ByVal codeEditorVbClass As IRichEditControl)
            Me.codeEditorCs = codeEditorCs
            Me.codeEditorVb = codeEditorVb
            Me.codeEditorCsClass = codeEditorCsClass
            Me.codeEditorVbClass = codeEditorVbClass

            AddHandler codeEditorCs.InnerControl.InitializeDocument, AddressOf InitializeSyntaxHighlightForCs
            AddHandler codeEditorVb.InnerControl.InitializeDocument, AddressOf InitializeSyntaxHighlightForVb
            AddHandler codeEditorCsClass.InnerControl.InitializeDocument, AddressOf InitializeSyntaxHighlightForCsClass
            AddHandler codeEditorVbClass.InnerControl.InitializeDocument, AddressOf InitializeSyntaxHighlightForVBClass
        End Sub


        Public ReadOnly Property CurrentCodeEditor() As InnerRichEditControl
            Get
                If CurrentExampleLanguage = ExampleLanguage.Csharp Then
                    Return codeEditorCs.InnerControl
                Else
                    Return codeEditorVb.InnerControl
                End If
            End Get
        End Property

        Public ReadOnly Property CurrentCodeClassEditor() As InnerRichEditControl
            Get
                If CurrentExampleLanguage = ExampleLanguage.Csharp Then
                    Return codeEditorCsClass.InnerControl
                Else
                    Return codeEditorVbClass.InnerControl
                End If

            End Get
        End Property
        Public ReadOnly Property LastExampleCodeModifiedTime() As Date
            Get
                Return lastExampleCodeModifiedTime_Renamed
            End Get
        End Property

        Public ReadOnly Property RichEditTextChanged() As Boolean
            Get
                Return richEditTextChanged_Renamed
            End Get
        End Property


        Public Property CurrentExampleLanguage() As ExampleLanguage
            Get
                Return current
            End Get
            Set(ByVal value As ExampleLanguage)
                Try
                    UnsubscribeRichEditEvents()
                    current = value
                Finally
                    SubscribeRichEditEvent()
                    forceTextChangesCounter = 0 ' no changes in that richEdit (CurrentCodeEditor)
                    richEditTextChanged_Renamed = True
                End Try
            End Set
        End Property
        Private Sub richEditControl_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
            If forceTextChangesCounter <= 0 Then
                richEditTextChanged_Renamed = True
                lastExampleCodeModifiedTime_Renamed = Date.Now
            Else
                forceTextChangesCounter -= 1
            End If
        End Sub

        Public Function ShowExample(ByVal oldExample As CodeExample, ByVal newExample As CodeExample) As String
            Dim richEditControlCs As InnerRichEditControl = codeEditorCs.InnerControl
            Dim richEditControlVb As InnerRichEditControl = codeEditorVb.InnerControl
            Dim richEditControlCsClass As InnerRichEditControl = codeEditorCsClass.InnerControl
            Dim richEditControlVbClass As InnerRichEditControl = codeEditorVbClass.InnerControl

            If oldExample IsNot Nothing Then
                'save edited example
                oldExample.CodeCS = richEditControlCs.Text
                oldExample.CodeVB = richEditControlVb.Text
                oldExample.CodeCsHelper = richEditControlCsClass.Text
                oldExample.CodeVbHelper = richEditControlVbClass.Text
            End If
            Dim exampleCode As String = String.Empty
            If newExample IsNot Nothing Then
                Try
                    forceTextChangesCounter = 2
                    exampleCode = If(CurrentExampleLanguage = ExampleLanguage.Csharp, newExample.CodeCS, newExample.CodeVB)
                    richEditControlCs.Text = newExample.CodeCS
                    richEditControlVb.Text = newExample.CodeVB
                    richEditControlCsClass.Text = newExample.CodeCsHelper
                    richEditControlVbClass.Text = newExample.CodeVbHelper

                    richEditTextChanged_Renamed = False
                Finally
                    richEditTextChanged_Renamed = True
                End Try
            End If
            Return exampleCode
        End Function

        Private Sub UpdatePageBackground(ByVal codeEvaluated As Boolean)
            CurrentCodeEditor.Document.SetPageBackground(If(codeEvaluated, DXColor.Empty, DXColor.FromArgb(&HFF, &HBC, &HC8)), True)
            CurrentCodeClassEditor.Document.SetPageBackground(If(codeEvaluated, DXColor.Empty, DXColor.FromArgb(&HFF, &HBC, &HC8)), True)
        End Sub

        Friend Sub BeforeCompile()
            UnsubscribeRichEditEvents()
        End Sub

        Friend Sub AfterCompile(ByVal codeExecutedWithoutExceptions As Boolean)
            UpdatePageBackground(codeExecutedWithoutExceptions)

            richEditTextChanged_Renamed = False
            ResetLastExampleModifiedTime()

            SubscribeRichEditEvent()
        End Sub
        Public Sub ResetLastExampleModifiedTime()
            lastExampleCodeModifiedTime_Renamed = Date.Now
        End Sub
        Private Sub UnsubscribeRichEditEvents()
            RemoveHandler CurrentCodeEditor.ContentChanged, AddressOf richEditControl_TextChanged
            RemoveHandler CurrentCodeClassEditor.ContentChanged, AddressOf richEditControl_TextChanged
        End Sub
        Private Sub SubscribeRichEditEvent()
            AddHandler CurrentCodeEditor.ContentChanged, AddressOf richEditControl_TextChanged
            AddHandler CurrentCodeClassEditor.ContentChanged, AddressOf richEditControl_TextChanged
        End Sub
        Private Sub InitializeSyntaxHighlightForCs(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorCs, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.Csharp))

            DisableRichEditFeatures(codeEditorCs)
        End Sub


        Private Sub InitializeSyntaxHighlightForVb(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorVb, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.VB))

            DisableRichEditFeatures(codeEditorVb)
        End Sub

        Private Sub InitializeSyntaxHighlightForCsClass(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorCsClass, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.Csharp))

            DisableRichEditFeatures(codeEditorCsClass)
        End Sub

        Private Sub InitializeSyntaxHighlightForVBClass(ByVal sender As Object, ByVal e As EventArgs)
            Dim syntaxHightlightInitializator As New SyntaxHightlightInitializeHelper()
            syntaxHightlightInitializator.Initialize(codeEditorVbClass, CodeExampleDemoUtils.GetCodeExampleFileExtension(ExampleLanguage.VB))

            DisableRichEditFeatures(codeEditorVbClass)
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

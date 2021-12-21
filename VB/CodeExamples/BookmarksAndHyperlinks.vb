Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module BookmarksAndHyperlinksActions

        Private Sub InsertBookmark(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertBookmark"
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            wordProcessor.BeginUpdate()
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim pos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Range.Start
            document.Bookmarks.Create(wordProcessor.Document.CreateRange(pos, 0), "Top")
            'Insert the hyperlink anchored to the created bookmark:
            Dim pos1 As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition((wordProcessor.Document.Range.[End]).ToInt() + 25)
            document.Hyperlinks.Create(wordProcessor.Document.InsertText(pos1, "get to the top"))
            document.Hyperlinks(CInt((0))).Anchor = "Top"
            wordProcessor.EndUpdate()
#End Region  ' #InsertBookmark
        End Sub

        Private Sub InsertHyperlink(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertHyperlink"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim hPos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = wordProcessor.Document.Range.Start
            document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"))
            document.Hyperlinks(CInt((0))).NavigateUri = "https://www.devexpress.com/Products/NET/Controls/WinForms/Rich_Editor/"
            document.Hyperlinks(CInt((0))).ToolTip = "WinForms Rich Text Editor"
#End Region  ' #InsertHyperlink
        End Sub
    End Module
End Namespace

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module BookmarksAndHyperlinksActions

        Public InsertBookmarkAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BookmarksAndHyperlinksActions.InsertBookmark

        Public InsertHyperlinkAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.BookmarksAndHyperlinksActions.InsertHyperlink

        Private Sub InsertBookmark(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertBookmark"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Access the start position of the document range.
            Dim pos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.Range.Start
            ' Create a bookmark at the document top.
            document.Bookmarks.Create(wordProcessor.Document.CreateRange(pos, 0), "Top")
            ' Create a hyperlink that navigates to the created bookmark.
            Dim pos1 As DevExpress.XtraRichEdit.API.Native.DocumentPosition = document.CreatePosition((wordProcessor.Document.Range.[End]).ToInt() + 25)
            document.Hyperlinks.Create(wordProcessor.Document.InsertText(pos1, "get to the top"))
            document.Hyperlinks(CInt((0))).Anchor = "Top"
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #InsertBookmark
        End Sub

        Private Sub InsertHyperlink(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertHyperlink"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the start position of the document range.
            Dim hPos As DevExpress.XtraRichEdit.API.Native.DocumentPosition = wordProcessor.Document.Range.Start
            ' Create a hyperlink at the specified position.
            document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"))
            ' Specify the URI to which the hyperlink navigates. 
            document.Hyperlinks(CInt((0))).NavigateUri = "https://devexpress.com"
            ' Specify the hyperlink tooltip.
            document.Hyperlinks(CInt((0))).ToolTip = "DevExpress"
#End Region  ' #InsertHyperlink
        End Sub
    End Module
End Namespace

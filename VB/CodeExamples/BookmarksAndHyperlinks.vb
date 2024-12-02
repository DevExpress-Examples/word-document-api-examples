Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module BookmarksAndHyperlinksActions

        Public InsertBookmarkAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf InsertBookmark

        Public InsertHyperlinkAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf InsertHyperlink

        Private Sub InsertBookmark(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertBookmark"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a bookmark at the start document position.
            document.Bookmarks.Create(document.CreateRange(document.Range.Start, 0), "Top")
            ' Create a hyperlink that navigates to the created bookmark.
            wordProcessor.Document.Paragraphs.Append()
            Dim hyperlinkRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = wordProcessor.Document.AppendText("get to the top")
            document.Hyperlinks.Create(hyperlinkRange)
            document.Hyperlinks(0).Anchor = "Top"
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #InsertBookmark
        End Sub

        Private Sub InsertHyperlink(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertHyperlink"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a hyperlink at the specified position.
            Dim hyperlinkRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.InsertText(document.Range.Start, "Follow me!")
            document.Hyperlinks.Create(hyperlinkRange)
            ' Specify the URI to which the hyperlink navigates. 
            document.Hyperlinks(0).NavigateUri = "https://devexpress.com"
            ' Specify the hyperlink tooltip.
            document.Hyperlinks(0).ToolTip = "DevExpress"
#End Region  ' #InsertHyperlink
        End Sub
    End Module
End Namespace

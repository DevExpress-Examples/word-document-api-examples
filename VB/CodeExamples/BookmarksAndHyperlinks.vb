Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Public NotInheritable Class BookmarksAndHyperlinksActions

        Private Sub New()
        End Sub

         Private Shared Sub InsertBookmark(ByVal server As RichEditDocumentServer)
'            #Region "#InsertBookmark"
            server.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            server.BeginUpdate()
            Dim document As Document = server.Document
            Dim pos As DocumentPosition = document.Range.Start
            document.Bookmarks.Create(server.Document.CreateRange(pos, 0), "Top")
            'Insert the hyperlink anchored to the created bookmark:
            Dim pos1 As DocumentPosition = document.CreatePosition((server.Document.Range.End).ToInt() + 25)
            document.Hyperlinks.Create(server.Document.InsertText(pos1, "get to the top"))
            document.Hyperlinks(0).Anchor = "Top"
            server.EndUpdate()
'            #End Region ' #InsertBookmark
         End Sub
         Private Shared Sub InsertHyperlink(ByVal server As RichEditDocumentServer)
'            #Region "#InsertHyperlink"
            Dim document As Document = server.Document
            Dim hPos As DocumentPosition = server.Document.Range.Start
            server.Document.Hyperlinks.Create(document.InsertText(hPos, "Follow me!"))
            document.Hyperlinks(0).NavigateUri = "https://www.devexpress.com/Products/NET/Controls/WinForms/Rich_Editor/"
            document.Hyperlinks(0).ToolTip = "WinForms Rich Text Editor"
'            #End Region ' #InsertHyperlink
         End Sub



    End Class
End Namespace

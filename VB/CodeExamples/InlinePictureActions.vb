Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class InlinePicturesActions

        Public Shared ImageCollectionAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.InlinePicturesActions.ImageCollection

        Public Shared SaveImageToFileAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.InlinePicturesActions.SaveImageToFile

        Private Shared Sub ImageCollection(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ImageCollection"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Obtain all images contained in the document.
            Dim images As DevExpress.XtraRichEdit.API.Native.ReadOnlyDocumentImageCollection = document.Images
            ' If the image width exceeds 50 millimeters, 
            ' scale the image proportionally to half its size.
            For i As Integer = 0 To images.Count - 1
                If images(CInt((i))).Size.Width > DevExpress.Office.Utils.Units.MillimetersToDocumentsF(50) Then
                    images(CInt((i))).ScaleX /= 2
                    images(CInt((i))).ScaleY /= 2
                End If
            Next
#End Region  ' #ImageCollection
        End Sub

        Private Shared Sub SaveImageToFile(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SaveImageToFile"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a document range.
            Dim myRange As DevExpress.XtraRichEdit.API.Native.DocumentRange = document.CreateRange(0, 100)
            ' Obtain all images in the target range.
            Dim images As DevExpress.XtraRichEdit.API.Native.ReadOnlyDocumentImageCollection = document.Images.[Get](myRange)
            If images.Count > 0 Then
                ' Save the first retrieved image as a PNG file.
                Dim myImage As DevExpress.Office.Utils.OfficeImage = images(CInt((0))).Image
                Dim image As System.Drawing.Image = myImage.NativeImage
                Dim imageName As String = System.[String].Format("Image_at_pos_{0}.png", images(CInt((0))).Range.Start.ToInt())
                image.Save(imageName)
                ' Open the File Explorer and select the saved image.
                System.Diagnostics.Process.Start("explorer.exe", "/select," & imageName)
            End If
#End Region  ' #SaveImageToFile
        End Sub
    End Class
End Namespace

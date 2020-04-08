Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Friend Class InlinePicturesActions
		Private Shared Sub ImageFromFile(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ImageFromFile"
			Dim document As Document = wordProcessor.Document
			Dim pos As DocumentPosition = document.Range.Start
			document.Images.Insert(pos, DocumentImageSource.FromFile("Documents\beverages.png"))
'			#End Region ' #ImageFromFile
		End Sub

		Private Shared Sub ImageCollection(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ImageCollection"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim images As ReadOnlyDocumentImageCollection = document.Images
			' If the width of an image exceeds 50 millimeters, 
			' the image is scaled proportionally to half its size.
			For i As Integer = 0 To images.Count - 1
				If images(i).Size.Width > DevExpress.Office.Utils.Units.MillimetersToDocumentsF(50) Then
'INSTANT VB WARNING: Instant VB cannot determine whether both operands of this division are integer types - if they are then you should use the VB integer division operator:
					images(i).ScaleX /= 2
'INSTANT VB WARNING: Instant VB cannot determine whether both operands of this division are integer types - if they are then you should use the VB integer division operator:
					images(i).ScaleY /= 2
				End If
			Next i
'			#End Region ' #ImageCollection
		End Sub

		Private Shared Sub SaveImageToFile(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#SaveImageToFile"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\MovieRentals.docx", DocumentFormat.OpenXml)
			Dim myRange As DocumentRange = document.CreateRange(0, 100)
			Dim images As ReadOnlyDocumentImageCollection = document.Images.Get(myRange)
			If images.Count > 0 Then
				Dim myImage As DevExpress.Office.Utils.OfficeImage = images(0).Image
				Dim image As System.Drawing.Image = myImage.NativeImage
				Dim imageName As String = String.Format("Image_at_pos_{0}.png", images(0).Range.Start.ToInt())
				image.Save(imageName)
				System.Diagnostics.Process.Start("explorer.exe", "/select," & imageName)
			End If
'			#End Region ' #SaveImageToFile
		End Sub
	End Class
End Namespace


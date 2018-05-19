Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class ShapesActions

        Private Shared Sub AddFloatingPicture(ByVal server As RichEditDocumentServer)
'            #Region "#AddFloatingPicture"
            Dim document As Document = server.Document
            document.AppendText("Line One" & ControlChars.Lf & "Line Two" & ControlChars.Lf & "Line Three")
            Dim myPicture As Shape = document.Shapes.InsertPicture(document.CreatePosition(15), System.Drawing.Image.FromFile("Documents\beverages.png"))
            myPicture.HorizontalAlignment = ShapeHorizontalAlignment.Center
'            #End Region ' #AddFloatingPicture
        End Sub

        Private Shared Sub FloatingPictureOffset(ByVal server As RichEditDocumentServer)
'            #Region "#FloatingPictureOffset"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            document.Unit = DevExpress.Office.DocumentUnit.Centimeter
            Dim myPicture As Shape = document.Shapes(1)
            ' Clear the qualitative positioning to allow positioning by specifying the numerical offset. 
            myPicture.HorizontalAlignment = ShapeHorizontalAlignment.None
            myPicture.VerticalAlignment = ShapeVerticalAlignment.None
            ' Specify the reference item for positioning.
            myPicture.RelativeHorizontalPosition = ShapeRelativeHorizontalPosition.LeftMargin
            myPicture.RelativeVerticalPosition = ShapeRelativeVerticalPosition.TopMargin
            ' Specify the offset value.
            myPicture.Offset = New System.Drawing.PointF(4.5F, 2.0F)
'            #End Region ' #FloatingPictureOffset
        End Sub

        Private Shared Sub ChangeZorderAndWrapping(ByVal server As RichEditDocumentServer)
'            #Region "#ChangeZorderAndWrapping"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            Dim myPicture As Shape = document.Shapes(1)
            myPicture.VerticalAlignment = ShapeVerticalAlignment.Top
            myPicture.ZOrder = document.Shapes(0).ZOrder - 1
            myPicture.TextWrapping = TextWrappingType.BehindText
'            #End Region ' #ChangeZorderAndWrapping
        End Sub

        Private Shared Sub AddTextBox(ByVal server As RichEditDocumentServer)
'            #Region "#AddTextBox"
            Dim document As Document = server.Document
            document.AppendText("Line One" & ControlChars.Lf & "Line Two" & ControlChars.Lf & "Line Three")
            Dim myTextBox As Shape = document.Shapes.InsertTextBox(document.CreatePosition(15))
            myTextBox.HorizontalAlignment = ShapeHorizontalAlignment.Center
            ' Specify the text box background color.
            myTextBox.Fill.Color = System.Drawing.Color.WhiteSmoke
            ' Draw a border around the text box.
            myTextBox.Line.Color = System.Drawing.Color.Black
            myTextBox.Line.Thickness = 1
            ' Modify text box content.
            Dim textBoxDocument As SubDocument = myTextBox.TextBox.Document
            textBoxDocument.AppendText("TextBox Text")
            Dim cp As CharacterProperties = textBoxDocument.BeginUpdateCharacters(textBoxDocument.Range.Start, 7)
            cp.ForeColor = System.Drawing.Color.Orange
            cp.FontSize = 24
            textBoxDocument.EndUpdateCharacters(cp)
'            #End Region ' #AddTextBox
        End Sub

        Private Shared Sub InsertRichTextInTextBox(ByVal server As RichEditDocumentServer)
'            #Region "#InsertRichTextInTextBox"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            Dim myTextBox As Shape = document.Shapes(0)
            ' Allow text box resize to fit contents.
            myTextBox.TextBox.HeightRule = TextBoxSizeRule.Auto
            Dim boxedDocument As SubDocument = myTextBox.TextBox.Document
            Dim appendPosition As Integer = myTextBox.TextBox.Document.Range.End.ToInt()
            ' Append the second paragraph of the main document to the boxed text.
            Dim newRange As DocumentRange = boxedDocument.AppendDocumentContent(document.Paragraphs(1).Range)
            boxedDocument.Paragraphs.Insert(newRange.Start)
            ' Insert an image form the main document into the text box.
            boxedDocument.Images.Insert(boxedDocument.CreatePosition(appendPosition), document.Images(0).Image.NativeImage)
            ' Resize the image so that its size equals the image in the main document.
            boxedDocument.Images(0).Size = document.Images(0).Size
'            #End Region ' #InsertRichTextInTextBox
        End Sub

        Private Shared Sub RotateAndResize(ByVal server As RichEditDocumentServer)
'            #Region "#RotateAndResize"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            For Each s As Shape In document.Shapes
                ' Rotate a text box and resize a floating picture.
                If s.TextBox Is Nothing Then
                    s.ScaleX = 0.1F
                    s.ScaleY = 0.1F
                Else
                    s.RotationAngle = 45
                End If
            Next s
'            #End Region ' #RotateAndResize
        End Sub

        Private Shared Sub SelectShape(ByVal server As RichEditDocumentServer)
'            #Region "#SelectShape"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
            document.Selection = document.Shapes(0).Range
'            #End Region ' #SelectShape
        End Sub
    End Class
End Namespace

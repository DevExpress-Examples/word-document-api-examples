Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class ShapesActions

        Public Shared AddFloatingPictureAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.AddFloatingPicture

        Public Shared FloatingPictureOffsetAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.FloatingPictureOffset

        Public Shared ChangeZorderAndWrappingAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.ChangeZorderAndWrapping

        Public Shared AddTextBoxAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.AddTextBox

        Public Shared InsertRichTextInTextBoxAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.InsertRichTextInTextBox

        Public Shared RotateAndResizeAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ShapesActions.RotateAndResize

        Private Shared Sub AddFloatingPicture(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AddFloatingPicture"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Append text to the document.
            document.AppendText("Line One" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Two" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Three")
            ' Insert a picture at the specified position from the file. 
            Dim myPicture As DevExpress.XtraRichEdit.API.Native.Shape = document.Shapes.InsertPicture(document.CreatePosition(15), System.Drawing.Image.FromFile("Documents\beverages.png"))
            ' Specify the picture alignment.
            myPicture.HorizontalAlignment = DevExpress.XtraRichEdit.API.Native.ShapeHorizontalAlignment.Center
#End Region  ' #AddFloatingPicture
        End Sub

        Private Shared Sub FloatingPictureOffset(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#FloatingPictureOffset"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Specify the document's measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Centimeter
            If document.Shapes.Count > 1 Then
                ' Access a picture.
                Dim myPicture As DevExpress.XtraRichEdit.API.Native.Shape = document.Shapes(1)
                ' Clear the horizontal and vertical alignment values.
                myPicture.HorizontalAlignment = DevExpress.XtraRichEdit.API.Native.ShapeHorizontalAlignment.None
                myPicture.VerticalAlignment = DevExpress.XtraRichEdit.API.Native.ShapeVerticalAlignment.None
                ' The picture's horizontal position is relative to the left margin.
                myPicture.RelativeHorizontalPosition = DevExpress.XtraRichEdit.API.Native.ShapeRelativeHorizontalPosition.LeftMargin
                ' The picture's vertical position is relative to the top margin.
                myPicture.RelativeVerticalPosition = DevExpress.XtraRichEdit.API.Native.ShapeRelativeVerticalPosition.TopMargin
                ' Specify the offset value.
                myPicture.Offset = New System.Drawing.PointF(4.5F, 2.0F)
            End If
#End Region  ' #FloatingPictureOffset
        End Sub

        Private Shared Sub ChangeZorderAndWrapping(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeZorderAndWrapping"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.Shapes.Count > 1 Then
                ' Access a picture.
                Dim myPicture As DevExpress.XtraRichEdit.API.Native.Shape = document.Shapes(1)
                ' Align the picture vertically.
                myPicture.VerticalAlignment = DevExpress.XtraRichEdit.API.Native.ShapeVerticalAlignment.Top
                ' Specify the picture position in the z-order.
                myPicture.ZOrder = document.Shapes(CInt((0))).ZOrder - 1
                ' Display document text over the picture.
                myPicture.TextWrapping = DevExpress.XtraRichEdit.API.Native.TextWrappingType.BehindText
            End If
#End Region  ' #ChangeZorderAndWrapping
        End Sub

        Private Shared Sub AddTextBox(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#AddTextBox"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Append text to the document.
            document.AppendText("Line One" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Two" & Global.Microsoft.VisualBasic.Constants.vbLf & "Line Three")
            ' Insert a text box at the specified position.
            Dim myTextBox As DevExpress.XtraRichEdit.API.Native.Shape = document.Shapes.InsertTextBox(document.CreatePosition(15))
            ' Align the text box horizontally.
            myTextBox.HorizontalAlignment = DevExpress.XtraRichEdit.API.Native.ShapeHorizontalAlignment.Center
            ' Specify the text box background color.
            myTextBox.Fill.Color = System.Drawing.Color.WhiteSmoke
            ' Draw a border around the text box.
            myTextBox.Line.Color = System.Drawing.Color.Black
            myTextBox.Line.Thickness = 1
            ' Modify text box content.
            Dim textBoxDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = myTextBox.ShapeFormat.TextBox.Document
            textBoxDocument.AppendText("TextBox Text")
            ' Format the boxed text.
            Dim cp As DevExpress.XtraRichEdit.API.Native.CharacterProperties = textBoxDocument.BeginUpdateCharacters(textBoxDocument.Range.Start, 7)
            cp.ForeColor = System.Drawing.Color.Orange
            cp.FontSize = 24
            textBoxDocument.EndUpdateCharacters(cp)
#End Region  ' #AddTextBox
        End Sub

        Private Shared Sub InsertRichTextInTextBox(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#InsertRichTextInTextBox"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)

            ' Access a document.
            Dim document As Document = wordProcessor.Document

            ' Access a text box.
            Dim myTextBox As Shape = document.Shapes(0)

            ' Allow text box resize to fit contents.
            myTextBox.ShapeFormat.TextBox.HeightRule = TextBoxSizeRule.Auto
            Dim boxedDocument As SubDocument = myTextBox.ShapeFormat.TextBox.Document
            Dim appendPosition As Integer = myTextBox.ShapeFormat.TextBox.Document.Range.[End].ToInt()

            ' Append the second paragraph of the main document to the boxed text.
            Dim newRange As DocumentRange = boxedDocument.AppendDocumentContent(document.Paragraphs(CInt((1))).Range)
            boxedDocument.Paragraphs.Insert(newRange.Start)

            ' Insert an image form the main document into the text box.
            boxedDocument.Images.Insert(boxedDocument.CreatePosition(appendPosition), document.Images(CInt((0))).Image.NativeImage)

            ' Resize the image so that its size equals the image in the main document.
            boxedDocument.Images(0).Size = document.Images(0).Size
#End Region  ' #InsertRichTextInTextBox
        End Sub

        Private Shared Sub RotateAndResize(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#RotateAndResize"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check all shapes in the document.
            For Each s As DevExpress.XtraRichEdit.API.Native.Shape In document.Shapes
                ' Rotate pictures.
                If s.Type = DevExpress.XtraRichEdit.API.Native.ShapeType.Picture Then
                    ' Resize text boxes.
                    s.RotationAngle = 45
                Else
                    s.ScaleX = 0.1F
                    s.ScaleY = 0.1F
                End If
            Next
#End Region  ' #RotateAndResize
        End Sub
    End Class
End Namespace

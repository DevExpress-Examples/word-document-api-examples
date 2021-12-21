Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class WatermarkActions

        Public Shared CreateTextWatermarkAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.WatermarkActions.CreateTextWatermark

        Public Shared CreateImageWatermarkAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.WatermarkActions.CreateImageWatermark

        Private Shared Sub CreateTextWatermark(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateTextWatermark"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check whether the document sections have headers.
            For Each section As DevExpress.XtraRichEdit.API.Native.Section In document.Sections
                If Not section.HasHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.Primary) Then
                    ' Create an empty header.
                    Dim header As DevExpress.XtraRichEdit.API.Native.SubDocument = section.BeginUpdateHeader()
                    section.EndUpdateHeader(header)
                End If
            Next

            ' Specify text watermark options.
            Dim textWatermarkOptions As DevExpress.XtraRichEdit.API.Native.TextWatermarkOptions = New DevExpress.XtraRichEdit.API.Native.TextWatermarkOptions()
            textWatermarkOptions.Color = System.Drawing.Color.LightGray
            textWatermarkOptions.FontFamily = "Calibri"
            textWatermarkOptions.Layout = DevExpress.XtraRichEdit.API.Native.WatermarkLayout.Horizontal
            textWatermarkOptions.Semitransparent = True
            ' Add a text watermark to all document pages.
            document.WatermarkManager.SetText("CONFIDENTIAL", textWatermarkOptions)
#End Region  ' #CreateTextWatermark
        End Sub

        Private Shared Sub CreateImageWatermark(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateImageWatermark"
            'Check whether the document sections have headers.
            For Each section As DevExpress.XtraRichEdit.API.Native.Section In wordProcessor.Document.Sections
                If Not section.HasHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.Primary) Then
                    ' Create an empty header.
                    Dim header As DevExpress.XtraRichEdit.API.Native.SubDocument = section.BeginUpdateHeader()
                    section.EndUpdateHeader(header)
                End If
            Next

            ' Specify image watermark options.
            Dim imageWatermarkOptions As DevExpress.XtraRichEdit.API.Native.ImageWatermarkOptions = New DevExpress.XtraRichEdit.API.Native.ImageWatermarkOptions()
            imageWatermarkOptions.Washout = False
            imageWatermarkOptions.Scale = 2
            ' Add an image watermark to all document pages.
            wordProcessor.Document.WatermarkManager.SetImage(System.Drawing.Image.FromFile("Documents//DevExpress.png"), imageWatermarkOptions)
#End Region  ' #CreateImageWatermark
        End Sub
    End Class
End Namespace

Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
Imports System

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class HeadersAndFootersActions

        Public Shared CreateHeaderAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.HeadersAndFootersActions.CreateHeader

        Public Shared ModifyHeaderAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.HeadersAndFootersActions.ModifyHeader

        Private Shared Sub CreateHeader(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateHeader"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the first document section.
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Check whether the document already has a header (the same header for all pages).
            If Not firstSection.HasHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.Primary) Then
                ' Create a header.
                Dim newHeader As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader()
                newHeader.AppendText("Header")
                firstSection.EndUpdateHeader(newHeader)
            End If
#End Region  ' #CreateHeader
        End Sub

        Private Shared Sub ModifyHeader(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ModifyHeader"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Append a new section to the document.
            document.AppendSection()
            ' Access the first document section.
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Start to edit the header of the HeaderFooterType.First type.
            Dim myHeader As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.First)
            ' Change the header text.
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ")
            Dim fld As DevExpress.XtraRichEdit.API.Native.Field = myHeader.Fields.Create(range.[End], "PAGE \* ARABICDASH")
            ' Update all fields in the header.
            myHeader.Fields.Update()
            ' Finalize to edit the header.
            firstSection.EndUpdateHeader(myHeader)
            ' Display the header on the first document page.
            firstSection.DifferentFirstPage = True
#End Region  ' #ModifyHeader
        End Sub
    End Class
End Namespace

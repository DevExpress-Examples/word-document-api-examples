Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditEDocumentServerExample.CodeExamples

    Friend Class HeadersAndFootersActions

        Private Shared Sub CreateHeader(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateHeader"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Create an empty header.
            Dim newHeader As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader()
            firstSection.EndUpdateHeader(newHeader)
            ' Check whether the document already has a header (the same header for all pages).
            If firstSection.HasHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.Primary) Then
                Dim headerDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader()
                document.ChangeActiveDocument(headerDocument)
                document.CaretPosition = headerDocument.CreatePosition(0)
                firstSection.EndUpdateHeader(headerDocument)
            End If
#End Region  ' #CreateHeader
        End Sub

        Private Shared Sub ModifyHeader(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ModifyHeader"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.AppendSection()
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Modify the header of the HeaderFooterType.First type.
            Dim myHeader As DevExpress.XtraRichEdit.API.Native.SubDocument = firstSection.BeginUpdateHeader(DevExpress.XtraRichEdit.API.Native.HeaderFooterType.First)
            Dim range As DevExpress.XtraRichEdit.API.Native.DocumentRange = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ")
            Dim fld As DevExpress.XtraRichEdit.API.Native.Field = myHeader.Fields.Create(range.[End], "PAGE \* ARABICDASH")
            myHeader.Fields.Update()
            firstSection.EndUpdateHeader(myHeader)
            ' Display the header of the HeaderFooterType.First type on the first page.
            firstSection.DifferentFirstPage = True
#End Region  ' #ModifyHeader
        End Sub
    End Class
End Namespace

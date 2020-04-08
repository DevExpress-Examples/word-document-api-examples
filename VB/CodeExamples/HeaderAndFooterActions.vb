Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditEDocumentServerExample.CodeExamples
	Friend Class HeadersAndFootersActions

		Private Shared Sub CreateHeader(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CreateHeader"
			Dim document As Document = wordProcessor.Document
			Dim firstSection As Section = document.Sections(0)
			' Create an empty header.
			Dim newHeader As SubDocument = firstSection.BeginUpdateHeader()
			firstSection.EndUpdateHeader(newHeader)
			' Check whether the document already has a header (the same header for all pages).
			If firstSection.HasHeader(HeaderFooterType.Primary) Then
				Dim headerDocument As SubDocument = firstSection.BeginUpdateHeader()
				document.ChangeActiveDocument(headerDocument)
				document.CaretPosition = headerDocument.CreatePosition(0)
				firstSection.EndUpdateHeader(headerDocument)
			End If
'			#End Region ' #CreateHeader
		End Sub


		Private Shared Sub ModifyHeader(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ModifyHeader"
			Dim document As Document = wordProcessor.Document
			document.AppendSection()
			Dim firstSection As Section = document.Sections(0)
			' Modify the header of the HeaderFooterType.First type.
			Dim myHeader As SubDocument = firstSection.BeginUpdateHeader(HeaderFooterType.First)
			Dim range As DocumentRange = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE NUMBER ")
			Dim fld As Field = myHeader.Fields.Create(range.End, "PAGE \* ARABICDASH")
			myHeader.Fields.Update()
			firstSection.EndUpdateHeader(myHeader)
			' Display the header of the HeaderFooterType.First type on the first page.
			firstSection.DifferentFirstPage = True
'			#End Region ' #ModifyHeader
		End Sub
	End Class
End Namespace

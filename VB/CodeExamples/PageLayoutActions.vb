Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Friend Class PageLayoutActions
		Private Shared Sub LineNumbering(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#LineNumbering"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
			document.Unit = DevExpress.Office.DocumentUnit.Inch
			Dim sec As Section = document.Sections(0)
			sec.LineNumbering.CountBy = 2
			sec.LineNumbering.Start = 1
			sec.LineNumbering.Distance = 0.25F
			sec.LineNumbering.RestartType = LineNumberingRestart.NewSection
'			#End Region ' #LineNumbering
		End Sub

		Private Shared Sub CreateColumns(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CreateColumns"
			Dim document As Document = wordProcessor.Document
			document.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
			document.Unit = DevExpress.Office.DocumentUnit.Inch
			' Get the first section in a document
			Dim firstSection As Section = document.Sections(0)
			' Create columns and apply them to the document
			Dim sectionColumnsLayout As SectionColumnCollection = firstSection.Columns.CreateUniformColumns(firstSection.Page, 0.2F, 3)
			firstSection.Columns.SetColumns(sectionColumnsLayout)
'			#End Region ' #CreateColumns
		End Sub

		Private Shared Sub PrintLayout(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#PrintLayout"
			wordProcessor.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			Dim document As Document = wordProcessor.Document
			document.Unit = DevExpress.Office.DocumentUnit.Inch
			document.Sections(0).Page.PaperKind = System.Drawing.Printing.PaperKind.A6
			document.Sections(0).Page.Landscape = True
			document.Sections(0).Margins.Left = 2.0F
'			#End Region ' #PrintLayout
		End Sub

		Private Shared Sub TabStops(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#TabStops"
			Dim document As Document = wordProcessor.Document
			wordProcessor.LoadDocument("Documents\Grimm.docx", DocumentFormat.OpenXml)
			document.Unit = DevExpress.Office.DocumentUnit.Inch
			Dim tabs As TabInfoCollection = document.Paragraphs(0).BeginUpdateTabs(True)
			Dim tab1 As New TabInfo()
			' Sets tab stop at 2.5 inch
			tab1.Position = 2.5F
			tab1.Alignment = TabAlignmentType.Left
			tab1.Leader = TabLeaderType.MiddleDots
			tabs.Add(tab1)
			Dim tab2 As New TabInfo()
			tab2.Position = 5.5F
			tab2.Alignment = TabAlignmentType.Decimal
			tab2.Leader = TabLeaderType.EqualSign
			tabs.Add(tab2)
			document.Paragraphs(0).EndUpdateTabs(tabs)
'			#End Region ' #TabStops
		End Sub
	End Class
End Namespace

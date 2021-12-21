Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class PageLayoutActions

        Public Shared LineNumberingAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.PageLayoutActions.LineNumbering

        Public Shared CreateColumnsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.PageLayoutActions.CreateColumns

        Public Shared PrintLayoutAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.PageLayoutActions.PrintLayout

        Public Shared TabStopsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.PageLayoutActions.TabStops

        Private Shared Sub LineNumbering(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#LineNumbering"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Access the first document section.
            Dim sec As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Specify line numbering parameters for the section.
            sec.LineNumbering.CountBy = 2
            sec.LineNumbering.Start = 1
            sec.LineNumbering.Distance = 0.25F
            sec.LineNumbering.RestartType = DevExpress.XtraRichEdit.API.Native.LineNumberingRestart.NewSection
#End Region  ' #LineNumbering
        End Sub

        Private Shared Sub CreateColumns(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateColumns"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Access the first document section.
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            ' Create a uniform column layout. 
            Dim sectionColumnsLayout As DevExpress.XtraRichEdit.API.Native.SectionColumnCollection = firstSection.Columns.CreateUniformColumns(firstSection.Page, 0.2F, 3)
            ' Apply the column layout to the section.
            firstSection.Columns.SetColumns(sectionColumnsLayout)
#End Region  ' #CreateColumns
        End Sub

        Private Shared Sub PrintLayout(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#PrintLayout"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Specify page layout settings for the first document section.
            document.Sections(CInt((0))).Page.PaperKind = System.Drawing.Printing.PaperKind.A6
            document.Sections(CInt((0))).Page.Landscape = True
            document.Sections(CInt((0))).Margins.Left = 2.0F
#End Region  ' #PrintLayout
        End Sub

        Private Shared Sub TabStops(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#TabStops"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Start to modify tab stops in the first paragraph.
            Dim tabs As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = document.Paragraphs(CInt((0))).BeginUpdateTabs(True)
            ' Create the first tab stop.
            Dim tab1 As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            ' Specify the tab stop settings.
            tab1.Position = 2.5F
            tab1.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.Left
            tab1.Leader = DevExpress.XtraRichEdit.API.Native.TabLeaderType.MiddleDots
            ' Add the tab stop to the collection of tab stops.
            tabs.Add(tab1)
            ' Create the second tab stop.
            Dim tab2 As DevExpress.XtraRichEdit.API.Native.TabInfo = New DevExpress.XtraRichEdit.API.Native.TabInfo()
            ' Specify the tab stop settings.
            tab2.Position = 5.5F
            tab2.Alignment = DevExpress.XtraRichEdit.API.Native.TabAlignmentType.[Decimal]
            tab2.Leader = DevExpress.XtraRichEdit.API.Native.TabLeaderType.EqualSign
            ' Add the tab stop to the collection of tab stops.
            tabs.Add(tab2)
            ' Finalize to modify tab stops in a paragraph.
            document.Paragraphs(CInt((0))).EndUpdateTabs(tabs)
#End Region  ' #TabStops
        End Sub
    End Class
End Namespace

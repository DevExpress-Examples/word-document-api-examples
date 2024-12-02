Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class PageLayoutActions

        Public Shared LineNumberingAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf LineNumbering

        Public Shared CreateColumnsAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf CreateColumns

        Public Shared PrintLayoutAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf PrintLayout

        Public Shared TabStopsAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf TabStops

        Public Shared PageBordersAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf CreatePageBorders

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
            document.Sections(0).Page.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A6
            document.Sections(0).Page.Landscape = True
            document.Sections(0).Margins.Left = 2.0F
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
            Dim tabs As DevExpress.XtraRichEdit.API.Native.TabInfoCollection = document.Paragraphs(0).BeginUpdateTabs(True)
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
            document.Paragraphs(0).EndUpdateTabs(tabs)
#End Region  ' #TabStops
        End Sub

        Private Shared Sub CreatePageBorders(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreatePageBorders"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Generate a document with two sections and multiple pages in each section.
            document.AppendText(Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed)
            document.Paragraphs.Append()
            document.AppendSection()
            document.AppendText(Global.Microsoft.VisualBasic.Constants.vbFormFeed & Global.Microsoft.VisualBasic.Constants.vbFormFeed)
            Dim firstSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(0)
            Dim pageBorders1 As DevExpress.XtraRichEdit.API.Native.SectionPageBorders = firstSection.PageBorders
            ' Set page borders for the first page of the first section.
            SetPageBorders(pageBorders1.LeftBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            SetPageBorders(pageBorders1.TopBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            SetPageBorders(pageBorders1.RightBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            SetPageBorders(pageBorders1.BottomBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Single], 1F, System.Drawing.Color.Red)
            pageBorders1.AppliesTo = DevExpress.XtraRichEdit.API.Native.PageBorderAppliesTo.FirstPage
            Dim secondSection As DevExpress.XtraRichEdit.API.Native.Section = document.Sections(1)
            Dim pageBorders2 As DevExpress.XtraRichEdit.API.Native.SectionPageBorders = secondSection.PageBorders
            ' Set page borders for all pages of the second section.
            SetPageBorders(pageBorders2.LeftBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            SetPageBorders(pageBorders2.TopBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            SetPageBorders(pageBorders2.RightBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            SetPageBorders(pageBorders2.BottomBorder, DevExpress.XtraRichEdit.API.Native.BorderLineStyle.[Double], 1.5F, System.Drawing.Color.Green)
            pageBorders2.AppliesTo = DevExpress.XtraRichEdit.API.Native.PageBorderAppliesTo.AllPages
            pageBorders2.ZOrder = DevExpress.XtraRichEdit.API.Native.PageBorderZOrder.Back
        End Sub

        Private Shared Sub SetPageBorders(ByVal border As DevExpress.XtraRichEdit.API.Native.PageBorder, ByVal lineStyle As DevExpress.XtraRichEdit.API.Native.BorderLineStyle, ByVal borderWidth As Single, ByVal color As System.Drawing.Color)
            border.LineStyle = lineStyle
            border.LineWidth = borderWidth
            border.LineColor = color
        End Sub
#End Region  ' #CreatePageBorders
    End Class
End Namespace

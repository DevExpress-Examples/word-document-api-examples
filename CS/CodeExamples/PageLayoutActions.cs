using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class PageLayoutActions
    {
        public static Action<RichEditDocumentServer> LineNumberingAction = LineNumbering;
        public static Action<RichEditDocumentServer> CreateColumnsAction = CreateColumns;
        public static Action<RichEditDocumentServer> PrintLayoutAction = PrintLayout;
        public static Action<RichEditDocumentServer> TabStopsAction = TabStops;
        public static Action<RichEditDocumentServer> PageBordersAction = CreatePageBorders;

        static void LineNumbering(RichEditDocumentServer wordProcessor)
        {
            #region #LineNumbering
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Access the first document section.
            Section sec = document.Sections[0];

            // Specify line numbering parameters for the section.
            sec.LineNumbering.CountBy = 2;
            sec.LineNumbering.Start = 1;
            sec.LineNumbering.Distance = 0.25f;
            sec.LineNumbering.RestartType = LineNumberingRestart.NewSection;
            #endregion #LineNumbering
        }

        static void CreateColumns(RichEditDocumentServer wordProcessor)
        {
            #region #CreateColumns
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Access the first document section.
            Section firstSection = document.Sections[0];

            // Create a uniform column layout. 
            SectionColumnCollection sectionColumnsLayout =
                firstSection.Columns.CreateUniformColumns(firstSection.Page, 0.2f, 3);

            // Apply the column layout to the section.
            firstSection.Columns.SetColumns(sectionColumnsLayout);
            #endregion #CreateColumns
        }

        static void PrintLayout(RichEditDocumentServer wordProcessor)
        {
            #region #PrintLayout
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Specify page layout settings for the first document section.
            document.Sections[0].Page.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A6;
            document.Sections[0].Page.Landscape = true;
            document.Sections[0].Margins.Left = 2.0f;
            #endregion #PrintLayout
        }

        static void TabStops(RichEditDocumentServer wordProcessor)
        {
            #region #TabStops
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Specify the document’s measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            // Start to modify tab stops in the first paragraph.
            TabInfoCollection tabs = document.Paragraphs[0].BeginUpdateTabs(true);

            // Create the first tab stop.
            TabInfo tab1 = new TabInfo();

            // Specify the tab stop settings.
            tab1.Position = 2.5f;
            tab1.Alignment = TabAlignmentType.Left;
            tab1.Leader = TabLeaderType.MiddleDots;

            // Add the tab stop to the collection of tab stops.
            tabs.Add(tab1);

            // Create the second tab stop.
            TabInfo tab2 = new TabInfo();

            // Specify the tab stop settings.
            tab2.Position = 5.5f;
            tab2.Alignment = TabAlignmentType.Decimal;
            tab2.Leader = TabLeaderType.EqualSign;

            // Add the tab stop to the collection of tab stops.
            tabs.Add(tab2);

            // Finalize to modify tab stops in a paragraph.
            document.Paragraphs[0].EndUpdateTabs(tabs);
            #endregion #TabStops
        }

        static void CreatePageBorders(RichEditDocumentServer wordProcessor)
        {
            #region #CreatePageBorders
            Document document = wordProcessor.Document;
            // Generate a document with two sections and multiple pages in each section.
            document.AppendText("\f\f\f");
            document.Paragraphs.Append();
            document.AppendSection();
            document.AppendText("\f\f");

            Section firstSection = document.Sections[0];
            SectionPageBorders pageBorders1 = firstSection.PageBorders;

            // Set page borders for the first page of the first section.
            SetPageBorders(pageBorders1.LeftBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            SetPageBorders(pageBorders1.TopBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            SetPageBorders(pageBorders1.RightBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            SetPageBorders(pageBorders1.BottomBorder, BorderLineStyle.Single, 1f, System.Drawing.Color.Red);
            pageBorders1.AppliesTo = PageBorderAppliesTo.FirstPage;

            Section secondSection = document.Sections[1];
            SectionPageBorders pageBorders2 = secondSection.PageBorders;

            // Set page borders for all pages of the second section.
            SetPageBorders(pageBorders2.LeftBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            SetPageBorders(pageBorders2.TopBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            SetPageBorders(pageBorders2.RightBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            SetPageBorders(pageBorders2.BottomBorder, BorderLineStyle.Double, 1.5f, System.Drawing.Color.Green);
            pageBorders2.AppliesTo = PageBorderAppliesTo.AllPages;
            pageBorders2.ZOrder = PageBorderZOrder.Back;
        }
        static void SetPageBorders(PageBorder border, BorderLineStyle lineStyle,
            float borderWidth, System.Drawing.Color color)
        {
            border.LineStyle = lineStyle;
            border.LineWidth = borderWidth;
            border.LineColor = color;
        }
        #endregion #CreatePageBorders
    }

}

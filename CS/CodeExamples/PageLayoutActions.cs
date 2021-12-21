using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class PageLayoutActions
    {
        public static Action<RichEditDocumentServer> LineNumberingAction = LineNumbering;
        public static Action<RichEditDocumentServer> CreateColumnsAction = CreateColumns;
        public static Action<RichEditDocumentServer> PrintLayoutAction = PrintLayout;
        public static Action<RichEditDocumentServer> TabStopsAction = TabStops;

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
            document.Sections[0].Page.PaperKind = System.Drawing.Printing.PaperKind.A6;
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
    }
}

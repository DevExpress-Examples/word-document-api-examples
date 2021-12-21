using System;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.IO;
using System.Drawing;
using DevExpress.Office.Utils;

namespace RichEditDocumentServerAPIExample.CodeExamples
{

    class TablesActions
    {
        public static Action<RichEditDocumentServer> CreateTableAction = CreateTable;
        public static Action<RichEditDocumentServer> CreateFixedTableAction = CreateFixedTable;
        public static Action<RichEditDocumentServer> ChangeTableColorAction = ChangeTableColor;
        public static Action<RichEditDocumentServer> CreateAndApplyTableStyleAction = CreateAndApplyTableStyle;
        public static Action<RichEditDocumentServer> UseConditionalStyleAction = UseConditionalStyle;
        public static Action<RichEditDocumentServer> ChangeColumnAppearanceAction = ChangeColumnAppearance;
        public static Action<RichEditDocumentServer> UseTableCellProcessorAction = UseTableCellProcessor;
        public static Action<RichEditDocumentServer> MergeCellsAction = MergeCells;
        public static Action<RichEditDocumentServer> SplitCellsAction = SplitCells;
        public static Action<RichEditDocumentServer> DeleteTableElementsAction = DeleteTableElements;
        public static Action<RichEditDocumentServer> WrapTextAroundTableAction = WrapTextAroundTable;


        static void CreateTable(RichEditDocumentServer wordProcessor)
        {

            #region #CreateTable
            // Access a document.
            Document document = wordProcessor.Document;

            // Insert a new table with one row and three columns at the document range's start position.
            Table tbl = document.Tables.Create(document.Range.Start, 1, 3, AutoFitBehaviorType.AutoFitToWindow);

            // Create a table header.
            document.InsertText(tbl[0, 0].Range.Start, "Name");
            document.InsertText(tbl[0, 1].Range.Start, "Size");
            document.InsertText(tbl[0, 2].Range.Start, "DateTime");

            // Insert table data.
            DirectoryInfo dirinfo = new DirectoryInfo("C:\\");
            try
            {
                // Start to modify the table.
                tbl.BeginUpdate();

                // Obtain the file list from the specified directory
                // and add each list item to the table as a row.
                foreach (FileInfo fi in dirinfo.GetFiles())
                {
                    TableRow row = tbl.Rows.Append();
                    TableCell cell = row.FirstCell;
                    string fileName = fi.Name;
                    string fileLength = String.Format("{0:N0}", fi.Length);
                    string fileLastTime = String.Format("{0:g}", fi.LastWriteTime);
                    document.InsertSingleLineText(cell.Range.Start, fileName);
                    document.InsertSingleLineText(cell.Next.Range.Start, fileLength);
                    document.InsertSingleLineText(cell.Next.Next.Range.Start, fileLastTime);
                }
                // Center the table header.
                foreach (Paragraph p in document.Paragraphs.Get(tbl.FirstRow.Range))
                {
                    p.Alignment = ParagraphAlignment.Center;
                }
            }
            finally
            {
                // Finalize to modify the table.
                tbl.EndUpdate();
            }
            #endregion #CreateTable
        }

        static void CreateFixedTable(RichEditDocumentServer wordProcessor)
        {
            #region #CreateFixedTable
            // Access a document.
            Document document = wordProcessor.Document;

            // Insert a new table with three rows and columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 3, 3);

            // Align the table.
            table.TableAlignment = TableRowAlignment.Center;

            // Set the table width to a fixed value.
            table.TableLayout = TableLayoutType.Fixed;
            table.PreferredWidthType = WidthType.Fixed;
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4f);

            // Set the second row height to a fixed value. 
            table.Rows[1].HeightType = HeightType.Exact;
            table.Rows[1].Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8f);

            // Set the cell width to a fixed value. 
            table[1, 1].PreferredWidthType = WidthType.Fixed;
            table[1, 1].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);

            // Merge table cells.
            table.MergeCells(table[1, 1], table[2, 1]);

            #endregion #CreateFixedTable
        }
        static void ChangeTableColor(RichEditDocumentServer wordProcessor)
        {
            #region #ChangeTableColor
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with three rows and five columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 3, 5, AutoFitBehaviorType.AutoFitToWindow);

            // Start to modify the table.
            table.BeginUpdate();

            // Specify the document's measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Millimeter;

            // Specify the amount of space between table cells.
            // The distance between cells is 4 mm.
            table.TableCellSpacing = 2;
            
            // Change the color of empty space between cells.
            table.TableBackgroundColor = Color.Violet;
            
            // Change the cell background color.
            table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellColor));
            table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellBorderColor));
            
            // Finalize to modify the table.
            table.EndUpdate();
            #endregion #ChangeTableColor

        }
        #region #@ChangeTableColor
        class TableHelper
        {
            public static void ChangeCellColor(TableCell cell, int i, int j)
            {
                cell.BackgroundColor = System.Drawing.Color.Yellow;
            }

            public static void ChangeCellBorderColor(TableCell cell, int i, int j)
            {
                cell.Borders.Bottom.LineColor = System.Drawing.Color.Red;
                cell.Borders.Left.LineColor = System.Drawing.Color.Red;
                cell.Borders.Right.LineColor = System.Drawing.Color.Red;
                cell.Borders.Top.LineColor = System.Drawing.Color.Red;
            }
        }
        #endregion #@ChangeTableColor
        static void CreateAndApplyTableStyle(RichEditDocumentServer wordProcessor)
        {
            #region #CreateAndApplyTableStyle
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create a new table style.
            TableStyle tStyleMain = document.TableStyles.CreateNew();

            // Specify table style options.
            tStyleMain.AllCaps = true;
            tStyleMain.FontName = "Segoe Condensed";
            tStyleMain.FontSize = 14;
            tStyleMain.Alignment = ParagraphAlignment.Center;
            tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Dotted;
            tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Dotted;
            tStyleMain.TableBorders.Top.LineThickness = 1.5f;
            tStyleMain.TableBorders.Top.LineStyle = TableBorderLineStyle.Double;
            tStyleMain.TableBorders.Left.LineThickness = 1.5f;
            tStyleMain.TableBorders.Left.LineStyle = TableBorderLineStyle.Double;
            tStyleMain.TableBorders.Bottom.LineThickness = 1.5f;
            tStyleMain.TableBorders.Bottom.LineStyle = TableBorderLineStyle.Double;
            tStyleMain.TableBorders.Right.LineThickness = 1.5f;
            tStyleMain.TableBorders.Right.LineStyle = TableBorderLineStyle.Double;
            tStyleMain.CellBackgroundColor = System.Drawing.Color.LightBlue;
            tStyleMain.TableLayout = TableLayoutType.Fixed;
            tStyleMain.Name = "MyTableStyle";
            
            // Add the style to the collection of styles.
            document.TableStyles.Add(tStyleMain);

            // Finalize to edit the document.
            document.EndUpdate();

            // Start to edit the document.
            document.BeginUpdate();

            // Create a table with three rows and columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 3, 3);

            // Set the table width to a fixed value.
            table.TableLayout = TableLayoutType.Fixed;
            table.PreferredWidthType = WidthType.Fixed;
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(3.5f);

            // Set the cell width to a fixed value. 
            table[1, 1].PreferredWidthType = WidthType.Fixed;
            table[1, 1].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            
            // Apply the created style to the table.
            table.Style = tStyleMain;
            
            // Finalize to edit the document.
            document.EndUpdate();

            // Insert text to the table cell.
            document.InsertText(table[1, 1].Range.Start, "STYLED");
            #endregion #CreateAndApplyTableStyle
        }

        static void UseConditionalStyle(RichEditDocumentServer wordProcessor)
        {
            #region #UseConditionalStyle
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\TableStyles.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create a new table style.
            TableStyle myNewStyle = document.TableStyles.CreateNew();

            // Specify the parent style.
            // The created style inherits from the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
            myNewStyle.Parent = document.TableStyles["Grid Table 5 Dark Accent 1"];

            // Create conditional styles for table elements.
            TableConditionalStyle myNewStyleForFirstRow =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow);
            myNewStyleForFirstRow.CellBackgroundColor = Color.PaleVioletRed;
            TableConditionalStyle myNewStyleForFirstColumn =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn);
            myNewStyleForFirstColumn.CellBackgroundColor = Color.PaleVioletRed;
            TableConditionalStyle myNewStyleForOddColumns =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding);
            myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(Color.PaleVioletRed);
            TableConditionalStyle myNewStyleForEvenColumns =
                myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding);
            myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(Color.PaleVioletRed);
            document.TableStyles.Add(myNewStyle);

            // Create a new table with four rows and columns at the document range's end position.
            Table table = document.Tables.Create(document.Range.End, 4, 4, AutoFitBehaviorType.AutoFitToWindow);
            
            // Apply the created style to the table.
            table.Style = myNewStyle;

            // Apply special formatting to the first row and first column.
            table.TableLook = TableLookTypes.ApplyFirstRow | TableLookTypes.ApplyFirstColumn;

            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #UseConditionalStyle
        }

        static void ChangeColumnAppearance(RichEditDocumentServer wordProcessor)
        {
            #region #ChangeColumnAppearance
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with three rows and ten columns.
            Table table = document.Tables.Create(document.Range.Start, 3, 10);

            // Start to modify the table.
            table.BeginUpdate();

            // Change cell background color and vertical alignment in the third column.
            table.ForEachRow(new TableRowProcessorDelegate(ChangeColumnAppearanceHelper.ChangeColumnColor));
            
            // Finalize to modify the table.
            table.EndUpdate();
            #endregion #ChangeColumnAppearance

        }
        #region #@ChangeColumnAppearance
        class ChangeColumnAppearanceHelper
        {
            public static void ChangeColumnColor(TableRow row, int rowIndex)
            {
                row[2].BackgroundColor = System.Drawing.Color.LightCyan;
                row[2].VerticalAlignment = TableCellVerticalAlignment.Center;
            }
        }
        #endregion #@ChangeColumnAppearance

        static void UseTableCellProcessor(RichEditDocumentServer wordProcessor)
        {
            #region #UseTableCellProcessor
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with eight rows and columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 8, 8);

            // Start to modify the table.
            table.BeginUpdate();

            // Use the TableCellProcessorDelegate delegate to pass each table cell 
            // to the method that miltiplies numbers and output the result cells.
            table.ForEachCell(new TableCellProcessorDelegate(UseTableCellProcessorHelper.MakeMultiplicationCell));
            
            // Finalize to modify the table.
            table.EndUpdate();
            #endregion #UseTableCellProcessor
        }
        #region #@UseTableCellProcessor
        class UseTableCellProcessorHelper
        {
            public static void MakeMultiplicationCell(TableCell cell, int i, int j)
            {
                SubDocument doc = cell.Range.BeginUpdateDocument();
                doc.InsertText(cell.Range.Start,
                    String.Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)));
                cell.Range.EndUpdateDocument(doc);
            }
        }
        #endregion #@UseTableCellProcessor

        static void MergeCells(RichEditDocumentServer wordProcessor)
        {
            #region #MergeCells
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with six rows and eight columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 6, 8);

            // Start to modify the table.
            table.BeginUpdate();

            // Merge cells vertically in the second column from the third to the sixth row.
            table.MergeCells(table[2, 1], table[5, 1]);

            // Merge cells horizontally in the third row from the fourth to the eighth column.
            table.MergeCells(table[2, 3], table[2, 7]);

            // Finalize to modify the table.
            table.EndUpdate();
            #endregion #MergeCells
        }
        static void SplitCells(RichEditDocumentServer wordProcessor)
        {
            #region #SplitCells
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with three rows and columns at the document range's start position.
            Table table = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.FixedColumnWidth, 350);
            
            // Split a cell vertically to three cells. 
            table.Cell(2, 1).Split(1, 3);
            #endregion #SplitCells
        }
        static void DeleteTableElements(RichEditDocumentServer wordProcessor)
        {
            #region #DeleteTableElements
            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with three rows and columns at the document range's start position.
            Table tbl = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToWindow);

            // Start to modify the table.
            tbl.BeginUpdate();
            
            // Delete the third table row.
            tbl.Rows[2].Delete();

            // Delete a cell.
            tbl.Cell(1, 1).Delete();

            // Finalize to modify the table.
            tbl.EndUpdate();
            #endregion #DeleteTableElements
        }

        static void WrapTextAroundTable(RichEditDocumentServer wordProcessor)
        {
            #region #WrapTextAroundTable
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Create a table with three rows and columns
            // at the start position of the fifth paragraph's range.
            Table table = document.Tables.Create(document.Paragraphs[4].Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToContents);

            // Start to modify the table.
            table.BeginUpdate();

            // Specifies the text wrapping type.
            table.TextWrappingType = TableTextWrappingType.Around;

            // Specify vertical alignment.
            table.RelativeVerticalPosition = TableRelativeVerticalPosition.Paragraph;
            table.VerticalAlignment = TableVerticalAlignment.None;
            table.OffsetYRelative = DevExpress.Office.Utils.Units.InchesToDocumentsF(2f);

            // Specify horizontal alignment.
            table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Margin;
            table.HorizontalAlignment = TableHorizontalAlignment.Center;

            // Set distance between the text and the table.
            table.MarginBottom = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginLeft = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginTop = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            table.MarginRight = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
            
            // Finalize to modify the table.
            table.EndUpdate();
            #endregion #WrapTextAroundTable
        }
    }
}

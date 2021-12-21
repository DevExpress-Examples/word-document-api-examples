Imports System
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.IO
Imports System.Drawing
Imports DevExpress.Office.Utils

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class TablesActions

        Public Shared CreateTableAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.CreateTable

        Public Shared CreateFixedTableAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.CreateFixedTable

        Public Shared ChangeTableColorAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.ChangeTableColor

        Public Shared CreateAndApplyTableStyleAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.CreateAndApplyTableStyle

        Public Shared UseConditionalStyleAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.UseConditionalStyle

        Public Shared ChangeColumnAppearanceAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.ChangeColumnAppearance

        Public Shared UseTableCellProcessorAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.UseTableCellProcessor

        Public Shared MergeCellsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.MergeCells

        Public Shared SplitCellsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.SplitCells

        Public Shared DeleteTableElementsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.DeleteTableElements

        Public Shared WrapTextAroundTableAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.WrapTextAroundTable

        Private Shared Sub CreateTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateTable"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Insert a new table with one row and three columns at the document range's start position.
            Dim tbl As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 1, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            ' Create a table header.
            document.InsertText(tbl(CInt((0)), CInt((0))).Range.Start, "Name")
            document.InsertText(tbl(CInt((0)), CInt((1))).Range.Start, "Size")
            document.InsertText(tbl(CInt((0)), CInt((2))).Range.Start, "DateTime")
            ' Insert table data.
            Dim dirinfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo("C:\")
            Try
                ' Start to modify the table.
                tbl.BeginUpdate()
                ' Obtain the file list from the specified directory
                ' and add each list item to the table as a row.
                For Each fi As System.IO.FileInfo In dirinfo.GetFiles()
                    Dim row As DevExpress.XtraRichEdit.API.Native.TableRow = tbl.Rows.Append()
                    Dim cell As DevExpress.XtraRichEdit.API.Native.TableCell = row.FirstCell
                    Dim fileName As String = fi.Name
                    Dim fileLength As String = System.[String].Format("{0:N0}", fi.Length)
                    Dim fileLastTime As String = System.[String].Format("{0:g}", fi.LastWriteTime)
                    document.InsertSingleLineText(cell.Range.Start, fileName)
                    document.InsertSingleLineText(cell.[Next].Range.Start, fileLength)
                    document.InsertSingleLineText(cell.[Next].[Next].Range.Start, fileLastTime)
                Next

                ' Center the table header.
                For Each p As DevExpress.XtraRichEdit.API.Native.Paragraph In document.Paragraphs.[Get](tbl.FirstRow.Range)
                    p.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
                Next
            Finally
                ' Finalize to modify the table.
                tbl.EndUpdate()
            End Try
#End Region  ' #CreateTable
        End Sub

        Private Shared Sub CreateFixedTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateFixedTable"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Insert a new table with three rows and columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3)
            ' Align the table.
            table.TableAlignment = DevExpress.XtraRichEdit.API.Native.TableRowAlignment.Center
            ' Set the table width to a fixed value.
            table.TableLayout = DevExpress.XtraRichEdit.API.Native.TableLayoutType.Fixed
            table.PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4F)
            ' Set the second row height to a fixed value. 
            table.Rows(CInt((1))).HeightType = DevExpress.XtraRichEdit.API.Native.HeightType.Exact
            table.Rows(CInt((1))).Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8F)
            ' Set the cell width to a fixed value. 
            table(CInt((1)), CInt((1))).PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table(CInt((1)), CInt((1))).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            ' Merge table cells.
            table.MergeCells(table(1, 1), table(2, 1))
#End Region  ' #CreateFixedTable
        End Sub

        Private Shared Sub ChangeTableColor(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeTableColor"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with three rows and five columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 5, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            ' Start to modify the table.
            table.BeginUpdate()
            ' Specify the document's measure units.
            document.Unit = DevExpress.Office.DocumentUnit.Millimeter
            ' Specify the amount of space between table cells.
            ' The distance between cells is 4 mm.
            table.TableCellSpacing = 2
            ' Change the color of empty space between cells.
            table.TableBackgroundColor = System.Drawing.Color.Violet
            ' Change the cell background color.
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.TableHelper.ChangeCellColor))
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.TableHelper.ChangeCellBorderColor))
            ' Finalize to modify the table.
            table.EndUpdate()
#End Region  ' #ChangeTableColor
        End Sub

#Region "#@ChangeTableColor"
        Private Class TableHelper

            Public Shared Sub ChangeCellColor(ByVal cell As DevExpress.XtraRichEdit.API.Native.TableCell, ByVal i As Integer, ByVal j As Integer)
                cell.BackgroundColor = System.Drawing.Color.Yellow
            End Sub

            Public Shared Sub ChangeCellBorderColor(ByVal cell As DevExpress.XtraRichEdit.API.Native.TableCell, ByVal i As Integer, ByVal j As Integer)
                cell.Borders.Bottom.LineColor = System.Drawing.Color.Red
                cell.Borders.Left.LineColor = System.Drawing.Color.Red
                cell.Borders.Right.LineColor = System.Drawing.Color.Red
                cell.Borders.Top.LineColor = System.Drawing.Color.Red
            End Sub
        End Class

#End Region  ' #@ChangeTableColor
        Private Shared Sub CreateAndApplyTableStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateAndApplyTableStyle"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a new table style.
            Dim tStyleMain As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            ' Specify table style options.
            tStyleMain.AllCaps = True
            tStyleMain.FontName = "Segoe Condensed"
            tStyleMain.FontSize = 14
            tStyleMain.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Center
            tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.Dotted
            tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.Dotted
            tStyleMain.TableBorders.Top.LineThickness = 1.5F
            tStyleMain.TableBorders.Top.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.[Double]
            tStyleMain.TableBorders.Left.LineThickness = 1.5F
            tStyleMain.TableBorders.Left.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.[Double]
            tStyleMain.TableBorders.Bottom.LineThickness = 1.5F
            tStyleMain.TableBorders.Bottom.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.[Double]
            tStyleMain.TableBorders.Right.LineThickness = 1.5F
            tStyleMain.TableBorders.Right.LineStyle = DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle.[Double]
            tStyleMain.CellBackgroundColor = System.Drawing.Color.LightBlue
            tStyleMain.TableLayout = DevExpress.XtraRichEdit.API.Native.TableLayoutType.Fixed
            tStyleMain.Name = "MyTableStyle"
            ' Add the style to the collection of styles.
            document.TableStyles.Add(tStyleMain)
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a table with three rows and columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3)
            ' Set the table width to a fixed value.
            table.TableLayout = DevExpress.XtraRichEdit.API.Native.TableLayoutType.Fixed
            table.PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(3.5F)
            ' Set the cell width to a fixed value. 
            table(CInt((1)), CInt((1))).PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table(CInt((1)), CInt((1))).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            ' Apply the created style to the table.
            table.Style = tStyleMain
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Insert text to the table cell.
            document.InsertText(table(CInt((1)), CInt((1))).Range.Start, "STYLED")
#End Region  ' #CreateAndApplyTableStyle
        End Sub

        Private Shared Sub UseConditionalStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#UseConditionalStyle"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\TableStyles.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a new table style.
            Dim myNewStyle As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            ' Specify the parent style.
            ' The created style inherits from the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
            myNewStyle.Parent = document.TableStyles("Grid Table 5 Dark Accent 1")
            ' Create conditional styles for table elements.
            Dim myNewStyleForFirstRow As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.FirstRow)
            myNewStyleForFirstRow.CellBackgroundColor = System.Drawing.Color.PaleVioletRed
            Dim myNewStyleForFirstColumn As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.FirstColumn)
            myNewStyleForFirstColumn.CellBackgroundColor = System.Drawing.Color.PaleVioletRed
            Dim myNewStyleForOddColumns As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.OddColumnBanding)
            myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(System.Drawing.Color.PaleVioletRed)
            Dim myNewStyleForEvenColumns As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.EvenColumnBanding)
            myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(System.Drawing.Color.PaleVioletRed)
            document.TableStyles.Add(myNewStyle)
            ' Create a new table with four rows and columns at the document range's end position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.[End], 4, 4, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            ' Apply the created style to the table.
            table.Style = myNewStyle
            ' Apply special formatting to the first row and first column.
            table.TableLook = DevExpress.XtraRichEdit.API.Native.TableLookTypes.ApplyFirstRow Or DevExpress.XtraRichEdit.API.Native.TableLookTypes.ApplyFirstColumn
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #UseConditionalStyle
        End Sub

        Private Shared Sub ChangeColumnAppearance(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeColumnAppearance"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with three rows and ten columns.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 10)
            ' Start to modify the table.
            table.BeginUpdate()
            ' Change cell background color and vertical alignment in the third column.
            table.ForEachRow(New DevExpress.XtraRichEdit.API.Native.TableRowProcessorDelegate(AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.ChangeColumnAppearanceHelper.ChangeColumnColor))
            ' Finalize to modify the table.
            table.EndUpdate()
#End Region  ' #ChangeColumnAppearance
        End Sub

#Region "#@ChangeColumnAppearance"
        Private Class ChangeColumnAppearanceHelper

            Public Shared Sub ChangeColumnColor(ByVal row As DevExpress.XtraRichEdit.API.Native.TableRow, ByVal rowIndex As Integer)
                row(CInt((2))).BackgroundColor = System.Drawing.Color.LightCyan
                row(CInt((2))).VerticalAlignment = DevExpress.XtraRichEdit.API.Native.TableCellVerticalAlignment.Center
            End Sub
        End Class

#End Region  ' #@ChangeColumnAppearance
        Private Shared Sub UseTableCellProcessor(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#UseTableCellProcessor"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with eight rows and columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 8, 8)
            ' Start to modify the table.
            table.BeginUpdate()
            ' Use the TableCellProcessorDelegate delegate to pass each table cell 
            ' to the method that miltiplies numbers and output the result cells.
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RichEditDocumentServerAPIExample.CodeExamples.TablesActions.UseTableCellProcessorHelper.MakeMultiplicationCell))
            ' Finalize to modify the table.
            table.EndUpdate()
#End Region  ' #UseTableCellProcessor
        End Sub

#Region "#@UseTableCellProcessor"
        Private Class UseTableCellProcessorHelper

            Public Shared Sub MakeMultiplicationCell(ByVal cell As DevExpress.XtraRichEdit.API.Native.TableCell, ByVal i As Integer, ByVal j As Integer)
                Dim doc As DevExpress.XtraRichEdit.API.Native.SubDocument = cell.Range.BeginUpdateDocument()
                doc.InsertText(cell.Range.Start, System.[String].Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)))
                cell.Range.EndUpdateDocument(doc)
            End Sub
        End Class

#End Region  ' #@UseTableCellProcessor
        Private Shared Sub MergeCells(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#MergeCells"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with six rows and eight columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 6, 8)
            ' Start to modify the table.
            table.BeginUpdate()
            ' Merge cells vertically in the second column from the third to the sixth row.
            table.MergeCells(table(2, 1), table(5, 1))
            ' Merge cells horizontally in the third row from the fourth to the eighth column.
            table.MergeCells(table(2, 3), table(2, 7))
            ' Finalize to modify the table.
            table.EndUpdate()
#End Region  ' #MergeCells
        End Sub

        Private Shared Sub SplitCells(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SplitCells"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with three rows and columns at the document range's start position.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.FixedColumnWidth, 350)
            ' Split a cell vertically to three cells. 
            table.Cell(CInt((2)), CInt((1))).Split(1, 3)
#End Region  ' #SplitCells
        End Sub

        Private Shared Sub DeleteTableElements(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#DeleteTableElements"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with three rows and columns at the document range's start position.
            Dim tbl As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            ' Start to modify the table.
            tbl.BeginUpdate()
            ' Delete the third table row.
            tbl.Rows(CInt((2))).Delete()
            ' Delete a cell.
            tbl.Cell(CInt((1)), CInt((1))).Delete()
            ' Finalize to modify the table.
            tbl.EndUpdate()
#End Region  ' #DeleteTableElements
        End Sub

        Private Shared Sub WrapTextAroundTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#WrapTextAroundTable"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents\Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table with three rows and columns
            ' at the start position of the fifth paragraph's range.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Paragraphs(CInt((4))).Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToContents)
            ' Start to modify the table.
            table.BeginUpdate()
            ' Specifies the text wrapping type.
            table.TextWrappingType = DevExpress.XtraRichEdit.API.Native.TableTextWrappingType.Around
            ' Specify vertical alignment.
            table.RelativeVerticalPosition = DevExpress.XtraRichEdit.API.Native.TableRelativeVerticalPosition.Paragraph
            table.VerticalAlignment = DevExpress.XtraRichEdit.API.Native.TableVerticalAlignment.None
            table.OffsetYRelative = DevExpress.Office.Utils.Units.InchesToDocumentsF(2F)
            ' Specify horizontal alignment.
            table.RelativeHorizontalPosition = DevExpress.XtraRichEdit.API.Native.TableRelativeHorizontalPosition.Margin
            table.HorizontalAlignment = DevExpress.XtraRichEdit.API.Native.TableHorizontalAlignment.Center
            ' Set distance between the text and the table.
            table.MarginBottom = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginLeft = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginTop = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginRight = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            ' Finalize to modify the table.
            table.EndUpdate()
#End Region  ' #WrapTextAroundTable
        End Sub
    End Class
End Namespace

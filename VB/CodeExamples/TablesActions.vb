Imports System
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.IO
Imports System.Drawing
Imports DevExpress.Office.Utils

Namespace RTEDocumentServerExamples.CodeExamples

    Friend Class TablesActions

        Private Shared Sub CreateTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateTable"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Insert new table.
            Dim tbl As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 1, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            ' Create a table header.
            document.InsertText(tbl(CInt((0)), CInt((0))).Range.Start, "Name")
            document.InsertText(tbl(CInt((0)), CInt((1))).Range.Start, "Size")
            document.InsertText(tbl(CInt((0)), CInt((2))).Range.Start, "DateTime")
            ' Insert table data.
            Dim dirinfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo("C:\")
            Try
                tbl.BeginUpdate()
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
                tbl.EndUpdate()
            End Try

            tbl.Cell(CInt((2)), CInt((1))).Split(1, 3)
#End Region  ' #CreateTable
        End Sub

        Private Shared Sub CreateFixedTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateFixedTable"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3)
            table.TableAlignment = DevExpress.XtraRichEdit.API.Native.TableRowAlignment.Center
            table.TableLayout = DevExpress.XtraRichEdit.API.Native.TableLayoutType.Fixed
            table.PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4F)
            table.Rows(CInt((1))).HeightType = DevExpress.XtraRichEdit.API.Native.HeightType.Exact
            table.Rows(CInt((1))).Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8F)
            table(CInt((1)), CInt((1))).PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table(CInt((1)), CInt((1))).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            table.MergeCells(table(1, 1), table(2, 1))
#End Region  ' #CreateFixedTable
        End Sub

        Private Shared Sub ChangeTableColor(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeTableColor"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Create a table.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 5, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            table.BeginUpdate()
            ' Provide the space between table cells.
            ' The distance between cells will be 4 mm.
            document.Unit = DevExpress.Office.DocumentUnit.Millimeter
            table.TableCellSpacing = 2
            ' Change the color of empty space between cells.
            table.TableBackgroundColor = System.Drawing.Color.Violet
            'Change cell background color.
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RTEDocumentServerExamples.CodeExamples.TablesActions.TableHelper.ChangeCellColor))
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RTEDocumentServerExamples.CodeExamples.TablesActions.TableHelper.ChangeCellBorderColor))
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
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.BeginUpdate()
            ' Create a new table style.
            Dim tStyleMain As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            ' Specify style characteristics.
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
            'Add the style to the document.
            document.TableStyles.Add(tStyleMain)
            document.EndUpdate()
            document.BeginUpdate()
            ' Create a table.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3)
            table.TableLayout = DevExpress.XtraRichEdit.API.Native.TableLayoutType.Fixed
            table.PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4.5F)
            table(CInt((1)), CInt((1))).PreferredWidthType = DevExpress.XtraRichEdit.API.Native.WidthType.Fixed
            table(CInt((1)), CInt((1))).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            ' Apply a previously defined style.
            table.Style = tStyleMain
            document.EndUpdate()
            document.InsertText(table(CInt((1)), CInt((1))).Range.Start, "STYLED")
#End Region  ' #CreateAndApplyTableStyle
        End Sub

        Private Shared Sub UseConditionalStyle(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#UseConditionalStyle"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\TableStyles.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            document.BeginUpdate()
            ' Create a new style that is based on the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
            Dim myNewStyle As DevExpress.XtraRichEdit.API.Native.TableStyle = document.TableStyles.CreateNew()
            myNewStyle.Parent = document.TableStyles("Grid Table 5 Dark Accent 1")
            ' Create conditional styles (styles for table elements)
            Dim myNewStyleForFirstRow As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.FirstRow)
            myNewStyleForFirstRow.CellBackgroundColor = System.Drawing.Color.PaleVioletRed
            Dim myNewStyleForFirstColumn As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.FirstColumn)
            myNewStyleForFirstColumn.CellBackgroundColor = System.Drawing.Color.PaleVioletRed
            Dim myNewStyleForOddColumns As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.OddColumnBanding)
            myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(System.Drawing.Color.PaleVioletRed)
            Dim myNewStyleForEvenColumns As DevExpress.XtraRichEdit.API.Native.TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(DevExpress.XtraRichEdit.API.Native.ConditionalTableStyleFormattingTypes.EvenColumnBanding)
            myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(System.Drawing.Color.PaleVioletRed)
            document.TableStyles.Add(myNewStyle)
            ' Create a new table and apply a new style.
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.[End], 4, 4, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            table.Style = myNewStyle
            ' Specify which conditonal styles are in effect.
            table.TableLook = DevExpress.XtraRichEdit.API.Native.TableLookTypes.ApplyFirstRow Or DevExpress.XtraRichEdit.API.Native.TableLookTypes.ApplyFirstColumn
            document.EndUpdate()
#End Region  ' #UseConditionalStyle
        End Sub

        Private Shared Sub ChangeColumnAppearance(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ChangeColumnAppearance"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 10)
            table.BeginUpdate()
            'Change cell background color and vertical alignment in the third column.
            table.ForEachRow(New DevExpress.XtraRichEdit.API.Native.TableRowProcessorDelegate(AddressOf RTEDocumentServerExamples.CodeExamples.TablesActions.ChangeColumnAppearanceHelper.ChangeColumnColor))
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
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 8, 8)
            table.BeginUpdate()
            table.ForEachCell(New DevExpress.XtraRichEdit.API.Native.TableCellProcessorDelegate(AddressOf RTEDocumentServerExamples.CodeExamples.TablesActions.UseTableCellProcessorHelper.MakeMultiplicationCell))
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
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 6, 8)
            table.BeginUpdate()
            table.MergeCells(table(2, 1), table(5, 1))
            table.MergeCells(table(2, 3), table(2, 7))
            table.EndUpdate()
#End Region  ' #MergeCells
        End Sub

        Private Shared Sub SplitCells(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#SplitCells"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.FixedColumnWidth, 350)
            'split a cell to three: 
            table.Cell(CInt((2)), CInt((1))).Split(1, 3)
#End Region  ' #SplitCells
        End Sub

        Private Shared Sub DeleteTableElements(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#DeleteTableElements"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            Dim tbl As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToWindow)
            tbl.BeginUpdate()
            tbl.Rows(CInt((2))).Delete()
            tbl.Cell(CInt((1)), CInt((1))).Delete()
            tbl.EndUpdate()
#End Region  ' #DeleteTableElements
        End Sub

        Private Shared Sub WrapTextAroundTable(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#WrapTextAroundTable"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents//Grimm.docx")
            Dim table As DevExpress.XtraRichEdit.API.Native.Table = document.Tables.Create(document.Paragraphs(CInt((4))).Range.Start, 3, 3, DevExpress.XtraRichEdit.API.Native.AutoFitBehaviorType.AutoFitToContents)
            table.BeginUpdate()
            table.TextWrappingType = DevExpress.XtraRichEdit.API.Native.TableTextWrappingType.Around
            'Specify vertical alignment:
            table.RelativeVerticalPosition = DevExpress.XtraRichEdit.API.Native.TableRelativeVerticalPosition.Paragraph
            table.VerticalAlignment = DevExpress.XtraRichEdit.API.Native.TableVerticalAlignment.None
            table.OffsetYRelative = DevExpress.Office.Utils.Units.InchesToDocumentsF(2F)
            'Specify horizontal alignment:
            table.RelativeHorizontalPosition = DevExpress.XtraRichEdit.API.Native.TableRelativeHorizontalPosition.Margin
            table.HorizontalAlignment = DevExpress.XtraRichEdit.API.Native.TableHorizontalAlignment.Center
            'Set distance between the text and the table:
            table.MarginBottom = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginLeft = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginTop = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.MarginRight = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3F)
            table.EndUpdate()
#End Region  ' #WrapTextAroundTable
        End Sub
    End Class
End Namespace

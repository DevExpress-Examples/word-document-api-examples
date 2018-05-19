Imports System
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.IO
Imports System.Drawing

Namespace RTEDocumentServerExamples.CodeExamples

    Friend Class TablesAction
        Private Shared Sub CreateTable(ByVal server As RichEditDocumentServer)

'            #Region "#CreateTable"
            Dim document As Document = server.Document
            ' Insert new table.
            Dim tbl As Table = document.Tables.Create(document.Range.Start, 1, 3, AutoFitBehaviorType.AutoFitToWindow)
            ' Create a table header.
            document.InsertText(tbl(0, 0).Range.Start, "Name")
            document.InsertText(tbl(0, 1).Range.Start, "Size")
            document.InsertText(tbl(0, 2).Range.Start, "DateTime")
            ' Insert table data.
            Dim dirinfo As New DirectoryInfo("C:\")
            Try
                tbl.BeginUpdate()
                For Each fi As FileInfo In dirinfo.GetFiles()
                    Dim row As TableRow = tbl.Rows.Append()
                    Dim cell As TableCell = row.FirstCell
                    Dim fileName As String = fi.Name
                    Dim fileLength As String = String.Format("{0:N0}", fi.Length)
                    Dim fileLastTime As String = String.Format("{0:g}", fi.LastWriteTime)
                    document.InsertSingleLineText(cell.Range.Start, fileName)
                    document.InsertSingleLineText(cell.Next.Range.Start, fileLength)
                    document.InsertSingleLineText(cell.Next.Next.Range.Start, fileLastTime)
                Next fi
                ' Center the table header.
                For Each p As Paragraph In document.Paragraphs.Get(tbl.FirstRow.Range)
                    p.Alignment = ParagraphAlignment.Center
                Next p
            Finally
                tbl.EndUpdate()
            End Try
            tbl.Cell(2, 1).Split(1, 3)
'            #End Region ' #CreateTable
        End Sub

        Private Shared Sub CreateFixedTable(ByVal server As RichEditDocumentServer)
'            #Region "#CreateFixedTable"
            Dim document As Document = server.Document
            Dim table As Table = document.Tables.Create(document.Range.Start, 3, 3)

            table.TableAlignment = TableRowAlignment.Center
            table.TableLayout = TableLayoutType.Fixed
            table.PreferredWidthType = WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4F)

            table.Rows(1).HeightType = HeightType.Exact
            table.Rows(1).Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8F)

            table(1, 1).PreferredWidthType = WidthType.Fixed
            table(1, 1).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)

            table.MergeCells(table(1, 1), table(2, 1))

'            #End Region ' #CreateFixedTable
        End Sub
        Private Shared Sub ChangeTableColor(ByVal server As RichEditDocumentServer)
'            #Region "#ChangeTableColor"
            Dim document As Document = server.Document
            ' Create a table.
            Dim table As Table = document.Tables.Create(document.Range.Start, 3, 5, AutoFitBehaviorType.AutoFitToWindow)
            table.BeginUpdate()
            ' Provide the space between table cells.
            ' The distance between cells will be 4 mm.
            document.Unit = DevExpress.Office.DocumentUnit.Millimeter
            table.TableCellSpacing = 2
            ' Change the color of empty space between cells.
            table.TableBackgroundColor = Color.Violet
            'Change cell background color.
            table.ForEachCell(New TableCellProcessorDelegate(AddressOf TableHelper.ChangeCellColor))
            table.ForEachCell(New TableCellProcessorDelegate(AddressOf TableHelper.ChangeCellBorderColor))
            table.EndUpdate()
'            #End Region ' #ChangeTableColor

        End Sub
        #Region "#@ChangeTableColor"
        Private Class TableHelper
            Public Shared Sub ChangeCellColor(ByVal cell As TableCell, ByVal i As Integer, ByVal j As Integer)
                cell.BackgroundColor = System.Drawing.Color.Yellow
            End Sub

            Public Shared Sub ChangeCellBorderColor(ByVal cell As TableCell, ByVal i As Integer, ByVal j As Integer)
                cell.Borders.Bottom.LineColor = System.Drawing.Color.Red
                cell.Borders.Left.LineColor = System.Drawing.Color.Red
                cell.Borders.Right.LineColor = System.Drawing.Color.Red
                cell.Borders.Top.LineColor = System.Drawing.Color.Red
            End Sub
        End Class
        #End Region ' #@ChangeTableColor
        Private Shared Sub CreateAndApplyTableStyle(ByVal server As RichEditDocumentServer)
'            #Region "#CreateAndApplyTableStyle"
            Dim document As Document = server.Document
            document.BeginUpdate()
            ' Create a new table style.
            Dim tStyleMain As TableStyle = document.TableStyles.CreateNew()
            ' Specify style characteristics.
            tStyleMain.AllCaps = True
            tStyleMain.FontName = "Segoe Condensed"
            tStyleMain.FontSize = 14
            tStyleMain.Alignment = ParagraphAlignment.Center
            tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Dotted
            tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Dotted
            tStyleMain.TableBorders.Top.LineThickness = 1.5F
            tStyleMain.TableBorders.Top.LineStyle = TableBorderLineStyle.Double
            tStyleMain.TableBorders.Left.LineThickness = 1.5F
            tStyleMain.TableBorders.Left.LineStyle = TableBorderLineStyle.Double
            tStyleMain.TableBorders.Bottom.LineThickness = 1.5F
            tStyleMain.TableBorders.Bottom.LineStyle = TableBorderLineStyle.Double
            tStyleMain.TableBorders.Right.LineThickness = 1.5F
            tStyleMain.TableBorders.Right.LineStyle = TableBorderLineStyle.Double
            tStyleMain.CellBackgroundColor = System.Drawing.Color.LightBlue
            tStyleMain.TableLayout = TableLayoutType.Fixed
            tStyleMain.Name = "MyTableStyle"
            'Add the style to the document.
            document.TableStyles.Add(tStyleMain)
            document.EndUpdate()
            document.BeginUpdate()
            ' Create a table.
            Dim table As Table = document.Tables.Create(document.Range.Start, 3, 3)
            table.TableLayout = TableLayoutType.Fixed
            table.PreferredWidthType = WidthType.Fixed
            table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4.5F)
            table(1, 1).PreferredWidthType = WidthType.Fixed
            table(1, 1).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5F)
            ' Apply a previously defined style.
            table.Style = tStyleMain
            document.EndUpdate()

            document.InsertText(table(1, 1).Range.Start, "STYLED")
'            #End Region ' #CreateAndApplyTableStyle
        End Sub

        Private Shared Sub UseConditionalStyle(ByVal server As RichEditDocumentServer)
'            #Region "#UseConditionalStyle"
            Dim document As Document = server.Document
            document.LoadDocument("Documents\TableStyles.docx", DocumentFormat.OpenXml)
            document.BeginUpdate()

            ' Create a new style that is based on the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
            Dim myNewStyle As TableStyle = document.TableStyles.CreateNew()
            myNewStyle.Parent = document.TableStyles("Grid Table 5 Dark Accent 1")
            ' Create conditional styles (styles for table elements)
            Dim myNewStyleForFirstRow As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow)
            myNewStyleForFirstRow.CellBackgroundColor = Color.PaleVioletRed
            Dim myNewStyleForFirstColumn As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn)
            myNewStyleForFirstColumn.CellBackgroundColor = Color.PaleVioletRed
            Dim myNewStyleForOddColumns As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding)
            myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(Color.PaleVioletRed)
            Dim myNewStyleForEvenColumns As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding)
            myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(Color.PaleVioletRed)
            document.TableStyles.Add(myNewStyle)
            ' Create a new table and apply a new style.
            Dim table As Table = document.Tables.Create(document.Range.End, 4, 4, AutoFitBehaviorType.AutoFitToWindow)
            table.Style = myNewStyle
            ' Specify which conditonal styles are in effect.
            table.TableLook = TableLookTypes.ApplyFirstRow Or TableLookTypes.ApplyFirstColumn

            document.EndUpdate()
'            #End Region ' #UseConditionalStyle
        End Sub

        Private Shared Sub ChangeColumnAppearance(ByVal server As RichEditDocumentServer)
'            #Region "#ChangeColumnAppearance"
            Dim document As Document = server.Document
            Dim table As Table = document.Tables.Create(document.Range.Start, 3, 10)
            table.BeginUpdate()
            'Change cell background color and vertical alignment in the third column.
            table.ForEachRow(New TableRowProcessorDelegate(AddressOf ChangeColumnAppearanceHelper.ChangeColumnColor))
            table.EndUpdate()
'            #End Region ' #ChangeColumnAppearance

        End Sub
        #Region "#@ChangeColumnAppearance"
        Private Class ChangeColumnAppearanceHelper
            Public Shared Sub ChangeColumnColor(ByVal row As TableRow, ByVal rowIndex As Integer)
                row(2).BackgroundColor = System.Drawing.Color.LightCyan
                row(2).VerticalAlignment = TableCellVerticalAlignment.Center
            End Sub
        End Class
        #End Region ' #@ChangeColumnAppearance

        Private Shared Sub UseTableCellProcessor(ByVal server As RichEditDocumentServer)
'            #Region "#UseTableCellProcessor"
            Dim document As Document = server.Document
            Dim table As Table = document.Tables.Create(document.Range.Start, 8, 8)
            table.BeginUpdate()
            table.ForEachCell(New TableCellProcessorDelegate(AddressOf UseTableCellProcessorHelper.MakeMultiplicationCell))
            table.EndUpdate()
'            #End Region ' #UseTableCellProcessor
        End Sub
        #Region "#@UseTableCellProcessor"
        Private Class UseTableCellProcessorHelper
            Public Shared Sub MakeMultiplicationCell(ByVal cell As TableCell, ByVal i As Integer, ByVal j As Integer)
                Dim doc As SubDocument = cell.Range.BeginUpdateDocument()
                doc.InsertText(cell.Range.Start, String.Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)))
                cell.Range.EndUpdateDocument(doc)
            End Sub
        End Class
        #End Region ' #@UseTableCellProcessor

        Private Shared Sub MergeCells(ByVal server As RichEditDocumentServer)
'            #Region "#MergeCells"
            Dim document As Document = server.Document
            Dim table As Table = document.Tables.Create(document.Range.Start, 6, 8)
            table.BeginUpdate()
            table.MergeCells(table(2, 1), table(5, 1))
            table.MergeCells(table(2, 3), table(2, 7))
            table.EndUpdate()
'            #End Region ' #MergeCells
        End Sub
        Private Shared Sub SplitCells(ByVal server As RichEditDocumentServer)
'            #Region "#SplitCells"
            Dim document As Document = server.Document
            Dim table As Table = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.FixedColumnWidth, 350)
            'split a cell to three: 
            table.Cell(2, 1).Split(1, 3)
'            #End Region ' #SplitCells
        End Sub
        Private Shared Sub DeleteTableElements(ByVal server As RichEditDocumentServer)
'            #Region "#DeleteTableElements"
            Dim document As Document = server.Document
            Dim tbl As Table = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToWindow)
            tbl.BeginUpdate()
            tbl.Rows(2).Delete()
            tbl.Cell(1, 1).Delete()

            tbl.EndUpdate()
'            #End Region ' #DeleteTableElements
        End Sub
    End Class
End Namespace

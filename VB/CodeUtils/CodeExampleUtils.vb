Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports RichEditDocumentServerAPIExample.CodeExamples

Namespace RichEditDocumentServerAPIExample.CodeUtils

#Region "CodeExampleUtils"
    Public Module CodeExampleUtils

        Public Function InitData() As GroupsOfRichEditExamples
            Dim examples As GroupsOfRichEditExamples = New GroupsOfRichEditExamples()
#Region "GroupNodes"
            examples.Add(New RichEditNode("Basic Actions"))
            examples.Add(New RichEditNode("Bookmarks and Hyperlinks"))
            examples.Add(New RichEditNode("Comments Actions"))
            examples.Add(New RichEditNode("Custom Xml Actions"))
            examples.Add(New RichEditNode("Document Properties Actions"))
            examples.Add(New RichEditNode("Export Actions"))
            examples.Add(New RichEditNode("Field Actions"))
            examples.Add(New RichEditNode("Formatting Actions"))
            examples.Add(New RichEditNode("Form Fields Actions"))
            examples.Add(New RichEditNode("Header and Footer Actions"))
            examples.Add(New RichEditNode("Import Actions"))
            examples.Add(New RichEditNode("Inline Picture Actions"))
            examples.Add(New RichEditNode("Lists Actions"))
            examples.Add(New RichEditNode("Notes Actions"))
            examples.Add(New RichEditNode("Page Layout Actions"))
            examples.Add(New RichEditNode("Protection Actions"))
            examples.Add(New RichEditNode("Range Actions"))
            examples.Add(New RichEditNode("Shapes Actions"))
            examples.Add(New RichEditNode("Styles Actions"))
            examples.Add(New RichEditNode("Tables Actions"))
            examples.Add(New RichEditNode("Watermark Actions"))
#End Region
#Region "ExampleNodes"
            'Add nodes to the "Basic Actions" group of examples.
            examples(0).Groups.Add(New RichEditExample("Create a Document", String.Empty, String.Empty, CreateNewDocumentAction, True))
            examples(0).Groups.Add(New RichEditExample("Load a Document", String.Empty, String.Empty, LoadDocumentAction, True))
            examples(0).Groups.Add(New RichEditExample("Merge Documents", String.Empty, String.Empty, MergeDocumentsAction, True))
            examples(0).Groups.Add(New RichEditExample("Split a Document", String.Empty, String.Empty, SplitDocumentAction, False))
            examples(0).Groups.Add(New RichEditExample("Save a Document", String.Empty, String.Empty, SaveDocumentAction, False))
            examples(0).Groups.Add(New RichEditExample("Print a Document", String.Empty, String.Empty, PrintDocumentAction, False))
            'Add nodes to the "Bookmarks and Hyperlinks" group of examples.
            examples(1).Groups.Add(New RichEditExample("Insert a Bookmark", String.Empty, String.Empty, InsertBookmarkAction, True))
            examples(1).Groups.Add(New RichEditExample("Insert a Hyperlink", String.Empty, String.Empty, InsertHyperlinkAction, True))
            'Add nodes to the "Comments" group of examples.
            examples(2).Groups.Add(New RichEditExample("Create a Comment", String.Empty, String.Empty, CommentsActions.CreateCommentAction, True))
            examples(2).Groups.Add(New RichEditExample("Create a Nested Comment", String.Empty, String.Empty, CommentsActions.CreateNestedCommentAction, True))
            examples(2).Groups.Add(New RichEditExample("Delete a Comment", String.Empty, String.Empty, CommentsActions.DeleteCommentAction, True))
            examples(2).Groups.Add(New RichEditExample("Edit Comment Properties", String.Empty, String.Empty, CommentsActions.EditCommentPropertiesAction, True))
            examples(2).Groups.Add(New RichEditExample("Edit Comment Content", String.Empty, String.Empty, CommentsActions.EditCommentContentAction, True))
            'Add nodes to the "Custom XML parts" group of examples.
            examples(3).Groups.Add(New RichEditExample("Add a Custom Xml Part", String.Empty, String.Empty, CustomXmlActions.AddCustomXmlPartAction, True))
            examples(3).Groups.Add(New RichEditExample("Access a Custom Xml Part", String.Empty, String.Empty, CustomXmlActions.AccessCustomXmlPartAction, True))
            examples(3).Groups.Add(New RichEditExample("Remove a Custom Xml Part", String.Empty, String.Empty, CustomXmlActions.RemoveCustomXmlPartAction, True))
            'Add nodes to the "Document Properties" group of examples.
            examples(4).Groups.Add(New RichEditExample("Set Built-in Properties", String.Empty, String.Empty, StandardDocumentPropertiesAction, True))
            examples(4).Groups.Add(New RichEditExample("Set Custom Properties", String.Empty, String.Empty, CustomDocumentPropertiesAction, True))
            'Add nodes to the "Export" group of examples.
            examples(5).Groups.Add(New RichEditExample("Export a Range to HTML", String.Empty, String.Empty, ExportActions.ExportRangeToHtmlAction, False))
            examples(5).Groups.Add(New RichEditExample("Export a Range to Plain Text", String.Empty, String.Empty, ExportActions.ExportRangeToPlainTextAction, False))
            examples(5).Groups.Add(New RichEditExample("Convert DOCX to PDF", String.Empty, String.Empty, ExportActions.ExportToPDFAction, False))
            examples(5).Groups.Add(New RichEditExample("Convert HTML to PDF", String.Empty, String.Empty, ExportActions.ConvertHTMLtoPDFAction, False))
            examples(5).Groups.Add(New RichEditExample("Convert HTML to DOCX", String.Empty, String.Empty, ExportActions.ConvertHTMLtoDOCXAction, False))
            examples(5).Groups.Add(New RichEditExample("Convert DOCX to HTML", String.Empty, String.Empty, ExportActions.ExportToHTMLAction, False))
            examples(5).Groups.Add(New RichEditExample("Handle the Before Export Event", String.Empty, String.Empty, ExportActions.BeforeExportAction, False))
            'Add nodes to the "Fields" group of examples.
            examples(6).Groups.Add(New RichEditExample("Insert a Field", String.Empty, String.Empty, FieldActions.InsertFieldAction, True))
            examples(6).Groups.Add(New RichEditExample("Modify a Field", String.Empty, String.Empty, FieldActions.ModifyFieldCodeAction, True))
            examples(6).Groups.Add(New RichEditExample("Create a Field from a Range", String.Empty, String.Empty, FieldActions.CreateFieldFromRangeAction, True))
            'Add nodes to the "Formatting" group of examples.
            examples(7).Groups.Add(New RichEditExample("Format Text", String.Empty, String.Empty, FormattingActions.FormatTextAction, True))
            examples(7).Groups.Add(New RichEditExample("Change Spacing", String.Empty, String.Empty, FormattingActions.ChangeSpacingAction, True))
            examples(7).Groups.Add(New RichEditExample("Reset Character Formatting", String.Empty, String.Empty, FormattingActions.ResetCharacterFormattingAction, True))
            examples(7).Groups.Add(New RichEditExample("Format a Paragraph", String.Empty, String.Empty, FormattingActions.FormatParagraphAction, True))
            examples(7).Groups.Add(New RichEditExample("Reset Paragraph Formatting", String.Empty, String.Empty, FormattingActions.ResetParagraphFormattingAction, True))
            'Add nodes to the "Form Fields" group of examples.
            examples(8).Groups.Add(New RichEditExample("Insert a CheckBox", String.Empty, String.Empty, FormFieldsActions.InsertCheckBoxAction, True))
            'Add nodes to the "Headers and Footers" group of examples.
            examples(9).Groups.Add(New RichEditExample("Create a Header", String.Empty, String.Empty, HeadersAndFootersActions.CreateHeaderAction, True))
            examples(9).Groups.Add(New RichEditExample("Modify a Header", String.Empty, String.Empty, HeadersAndFootersActions.ModifyHeaderAction, True))
            'Add nodes to the "Import" group of examples.
            examples(10).Groups.Add(New RichEditExample("Import RTF Text", String.Empty, String.Empty, ImportActions.ImportRtfTextAction, True))
            examples(10).Groups.Add(New RichEditExample("Handle the Before Import Event", String.Empty, String.Empty, ImportActions.BeforeImportAction, True))
            'Add nodes to the "Inline Pictures" group of examples.
            examples(11).Groups.Add(New RichEditExample("Access an Image Collection", String.Empty, String.Empty, InlinePicturesActions.ImageCollectionAction, True))
            examples(11).Groups.Add(New RichEditExample("Save an Image to a File", String.Empty, String.Empty, InlinePicturesActions.SaveImageToFileAction, False))
            'Add nodes to the "Lists" group of examples.
            examples(12).Groups.Add(New RichEditExample("Create a Bulleted List", String.Empty, String.Empty, ListsActions.CreateBulletedListAction, True))
            examples(12).Groups.Add(New RichEditExample("Create a Numbered List", String.Empty, String.Empty, ListsActions.CreateNumberedListAction, True))
            examples(12).Groups.Add(New RichEditExample("Create a Multilevel List", String.Empty, String.Empty, ListsActions.CreateMultilevelListAction, True))
            'Add nodes to the "Notes" group of examples.
            examples(13).Groups.Add(New RichEditExample("Insert Footnotes", String.Empty, String.Empty, InsertFootnotesAction, True))
            examples(13).Groups.Add(New RichEditExample("Insert Endnotes", String.Empty, String.Empty, InsertEndnotesAction, True))
            examples(13).Groups.Add(New RichEditExample("Edit a Footnote", String.Empty, String.Empty, EditFootnoteAction, True))
            examples(13).Groups.Add(New RichEditExample("Edit an Endnote", String.Empty, String.Empty, EditEndnoteAction, True))
            examples(13).Groups.Add(New RichEditExample("Edit a Separator", String.Empty, String.Empty, EditSeparatorAction, True))
            examples(13).Groups.Add(New RichEditExample("Remove Notes", String.Empty, String.Empty, RemoveNotesAction, True))
            'Add nodes to the "Page Layout" group of examples.
            examples(14).Groups.Add(New RichEditExample("Add Line Numbering", String.Empty, String.Empty, PageLayoutActions.LineNumberingAction, True))
            examples(14).Groups.Add(New RichEditExample("Create Columns", String.Empty, String.Empty, PageLayoutActions.CreateColumnsAction, True))
            examples(14).Groups.Add(New RichEditExample("Adjust Page Layout", String.Empty, String.Empty, PageLayoutActions.PrintLayoutAction, True))
            examples(14).Groups.Add(New RichEditExample("Set Tab Stops", String.Empty, String.Empty, PageLayoutActions.TabStopsAction, True))
            'Add nodes to the "Protection" group of examples.
            examples(15).Groups.Add(New RichEditExample("Protect a Document", String.Empty, String.Empty, ProtectDocumentAction, False))
            examples(15).Groups.Add(New RichEditExample("Unprotect a Document", String.Empty, String.Empty, UnprotectDocumentAction, False))
            examples(15).Groups.Add(New RichEditExample("Create Range Permissions", String.Empty, String.Empty, CreateRangePermissionsAction, False))
            'Add nodes to the "Ranges" group of examples.
            examples(16).Groups.Add(New RichEditExample("Insert Text in a Range", String.Empty, String.Empty, RangeActions.InsertTextInRangeAction, True))
            examples(16).Groups.Add(New RichEditExample("Append Text to a Range", String.Empty, String.Empty, RangeActions.AppendTextToRangeAction, True))
            examples(16).Groups.Add(New RichEditExample("Append Text to a Paragraph", String.Empty, String.Empty, RangeActions.AppendToParagraphAction, True))
            'Add nodes to the "Shapes" group of examples.
            examples(17).Groups.Add(New RichEditExample("Add a Floating Picture", String.Empty, String.Empty, ShapesActions.AddFloatingPictureAction, True))
            examples(17).Groups.Add(New RichEditExample("Floating Picture Offset", String.Empty, String.Empty, ShapesActions.FloatingPictureOffsetAction, True))
            examples(17).Groups.Add(New RichEditExample("Change Z-Order and Wrapping", String.Empty, String.Empty, ShapesActions.ChangeZorderAndWrappingAction, True))
            examples(17).Groups.Add(New RichEditExample("Add a Text Box", String.Empty, String.Empty, ShapesActions.AddTextBoxAction, True))
            examples(17).Groups.Add(New RichEditExample("Insert Rich Text in a TextBox", String.Empty, String.Empty, ShapesActions.InsertRichTextInTextBoxAction, True))
            examples(17).Groups.Add(New RichEditExample("Rotate and Resize Shapes", String.Empty, String.Empty, ShapesActions.RotateAndResizeAction, True))
            'Add nodes to the "Styles" group of examples.
            examples(18).Groups.Add(New RichEditExample("Create a New Character Style", String.Empty, String.Empty, StylesAction.CreateNewCharacterStyleAction, True))
            examples(18).Groups.Add(New RichEditExample("Create a New Paragraph Style", String.Empty, String.Empty, StylesAction.CreateNewParagraphStyleAction, True))
            examples(18).Groups.Add(New RichEditExample("Create a New Linked Style", String.Empty, String.Empty, StylesAction.CreateNewLinkedStyleAction, False))
            'Add nodes to the "Tables" group of examples.
            examples(19).Groups.Add(New RichEditExample("Create a Table", String.Empty, String.Empty, TablesActions.CreateTableAction, True))
            examples(19).Groups.Add(New RichEditExample("Create a Fixed Table", String.Empty, String.Empty, TablesActions.CreateFixedTableAction, True))
            examples(19).Groups.Add(New RichEditExample("Change the Table Color", String.Empty, String.Empty, TablesActions.ChangeTableColorAction, True))
            examples(19).Groups.Add(New RichEditExample("Create and Apply a Table Style", String.Empty, String.Empty, TablesActions.CreateAndApplyTableStyleAction, True))
            examples(19).Groups.Add(New RichEditExample("Use a Conditional Style", String.Empty, String.Empty, TablesActions.UseConditionalStyleAction, True))
            examples(19).Groups.Add(New RichEditExample("Change Column Appearance", String.Empty, String.Empty, TablesActions.ChangeColumnAppearanceAction, True))
            examples(19).Groups.Add(New RichEditExample("Table Cell Processor", String.Empty, String.Empty, TablesActions.UseTableCellProcessorAction, True))
            examples(19).Groups.Add(New RichEditExample("Merge Cells", String.Empty, String.Empty, TablesActions.MergeCellsAction, True))
            examples(19).Groups.Add(New RichEditExample("Split Cells", String.Empty, String.Empty, TablesActions.SplitCellsAction, True))
            examples(19).Groups.Add(New RichEditExample("Delete Table Elements", String.Empty, String.Empty, TablesActions.DeleteTableElementsAction, True))
            examples(19).Groups.Add(New RichEditExample("Wrap Text Around a Table", String.Empty, String.Empty, TablesActions.WrapTextAroundTableAction, True))
            'Add nodes to the "Watermarks" group of examples.
            examples(20).Groups.Add(New RichEditExample("Create a Text Watermark", String.Empty, String.Empty, WatermarkActions.CreateTextWatermarkAction, True))
            examples(20).Groups.Add(New RichEditExample("Create an Image Watermark", String.Empty, String.Empty, WatermarkActions.CreateImageWatermarkAction, True))
            Return examples
#End Region
        End Function

        Public Function GatherExamplesFromProject(ByVal examplesPath As String, ByVal language As ExampleLanguage) As Dictionary(Of String, FileInfo)
            Dim result As Dictionary(Of String, FileInfo) = New Dictionary(Of String, FileInfo)()
            For Each fileName As String In Directory.GetFiles(examplesPath, "*" & GetCodeExampleFileExtension(language))
                result.Add(Path.GetFileNameWithoutExtension(fileName), New FileInfo(fileName))
            Next

            Return result
        End Function

        Public Function GetCodeExampleFileExtension(ByVal language As ExampleLanguage) As String
            If language = ExampleLanguage.VB Then Return ".vb"
            Return ".cs"
        End Function

        Public Function DeleteLeadingWhiteSpaces(ByVal lines As String(), ByVal stringToDelete As String) As String()
            Dim result As String() = New String(lines.Length - 1) {}
            Dim stringToDeleteLength As Integer = stringToDelete.Length
            For i As Integer = 0 To lines.Length - 1
                Dim index As Integer = lines(i).IndexOf(stringToDelete)
                result(i) = If(index >= 0, lines(i).Substring(index + stringToDeleteLength), lines(i))
            Next

            Return result
        End Function

        Public Function ConvertStringToHumanReadableForm(ByVal exampleName As String) As String
            Dim result As String = SplitCamelCase(exampleName)
            result = result.Replace(" In ", " in ")
            result = result.Replace(" And ", " and ")
            result = result.Replace(" To ", " to ")
            result = result.Replace(" From ", " from ")
            result = result.Replace(" With ", " with ")
            result = result.Replace(" By ", " by ")
            result = result.Replace(""""c, Microsoft.VisualBasic.Strings.ChrW(0))
            Return result
        End Function

        Private Function SplitCamelCase(ByVal exampleName As String) As String
            Dim length As Integer = exampleName.Length
            If length = 1 Then Return exampleName
            Dim result As StringBuilder = New StringBuilder(length * 2)
            For position As Integer = 0 To length - 1 - 1
                Dim current As Char = exampleName(position)
                Dim [next] As Char = exampleName(position + 1)
                result.Append(current)
                If Char.IsLower(current) AndAlso Char.IsUpper([next]) Then
                    result.Append(" "c)
                End If
            Next

            result.Append(exampleName(length - 1))
            Return result.ToString()
        End Function

        Public Function GetExamplePath(ByVal exampleFolderName As String) As String
            Dim examplesPath2 As String = Path.Combine(Directory.GetCurrentDirectory() & "\..\..\", exampleFolderName)
            If Directory.Exists(examplesPath2) Then Return examplesPath2
            Dim examplesPathInInsallation As String = GetRelativeDirectoryPath(exampleFolderName)
            Return examplesPathInInsallation
        End Function

        Public Function GetRelativeDirectoryPath(ByVal name As String) As String
            name = "Data\" & name
            Dim path As String = System.Windows.Forms.Application.StartupPath
            Dim s As String = "\"
            For i As Integer = 0 To 10
                If Directory.Exists(path & s & name) Then
                    Return path & s & name
                Else
                    s += "..\"
                End If
            Next

            Return ""
        End Function

        Public Function FindExamples(ByVal examples As Dictionary(Of String, FileInfo), ByVal exampleFinder As ExampleFinder) As GroupsOfRichEditExamples
            Dim richEditExamples As GroupsOfRichEditExamples = InitData()
            For Each sourceCodeItem As KeyValuePair(Of String, FileInfo) In examples
                Dim key As String = sourceCodeItem.Key
                Dim foundExamples As List(Of CodeExample) = exampleFinder.Process(examples(key))
                If foundExamples.Count = 0 Then Continue For
                For Each node As RichEditNode In richEditExamples
                    If Equals(node.Name, foundExamples(0).HumanReadableGroupName) AndAlso node.Groups.Count = foundExamples.Count Then
                        Dim i As Integer = 0
                        For Each example As RichEditExample In node.Groups
                            example.CodeCS = foundExamples(i).CodeCS
                            example.CodeVB = foundExamples(i).CodeVB
                            i += 1
                        Next
                    End If
                Next
            Next

            Return richEditExamples
        End Function

        Public Function DetectExampleLanguage(ByVal solutionFileNameWithoutExtenstion As String) As ExampleLanguage
            Dim projectPath As String = Directory.GetCurrentDirectory() & "\..\..\"
            Dim csproject As String() = Directory.GetFiles(projectPath, "*.csproj")
            If csproject.Length <> 0 AndAlso csproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".csproj") Then Return ExampleLanguage.Csharp
            Dim vbproject As String() = Directory.GetFiles(projectPath, "*.vbproj")
            If vbproject.Length <> 0 AndAlso vbproject(0).EndsWith(solutionFileNameWithoutExtenstion & ".vbproj") Then Return ExampleLanguage.VB
            Return ExampleLanguage.Csharp
        End Function
    End Module
#End Region
End Namespace

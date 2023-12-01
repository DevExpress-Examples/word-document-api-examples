using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using RichEditDocumentServerAPIExample.CodeExamples;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    #region CodeExampleUtils
    public static class CodeExampleUtils
    {
        public static GroupsOfRichEditExamples InitData()
        {
            GroupsOfRichEditExamples examples = new GroupsOfRichEditExamples();
            #region GroupNodes
            examples.Add(new RichEditNode("Basic Actions"));
            examples.Add(new RichEditNode("Bookmarks and Hyperlinks"));
            examples.Add(new RichEditNode("Comments Actions"));
            examples.Add(new RichEditNode("Content Controls Actions"));
            examples.Add(new RichEditNode("Custom Xml Actions"));
            examples.Add(new RichEditNode("Document Properties Actions"));
            examples.Add(new RichEditNode("Export Actions"));
            examples.Add(new RichEditNode("Field Actions"));
            examples.Add(new RichEditNode("Formatting Actions"));
            examples.Add(new RichEditNode("Form Fields Actions"));
            examples.Add(new RichEditNode("Header and Footer Actions"));
            examples.Add(new RichEditNode("Import Actions"));
            examples.Add(new RichEditNode("Inline Picture Actions"));
            examples.Add(new RichEditNode("Lists Actions"));
            examples.Add(new RichEditNode("Notes Actions"));
            examples.Add(new RichEditNode("Page Layout Actions"));
            examples.Add(new RichEditNode("Protection Actions"));
            examples.Add(new RichEditNode("Range Actions"));
            examples.Add(new RichEditNode("Shapes Actions"));
            examples.Add(new RichEditNode("Styles Actions"));
            examples.Add(new RichEditNode("Tables Actions"));
            examples.Add(new RichEditNode("Watermark Actions"));

            #endregion

            #region ExampleNodes
            //Add nodes to the "Basic Actions" group of examples.
            examples[0].Groups.Add(new RichEditExample("Create a Document", string.Empty, string.Empty, BasicActions.CreateNewDocumentAction, true));
            examples[0].Groups.Add(new RichEditExample("Load a Document", string.Empty, string.Empty, BasicActions.LoadDocumentAction, true));
            examples[0].Groups.Add(new RichEditExample("Merge Documents", string.Empty, string.Empty, BasicActions.MergeDocumentsAction, true));
            examples[0].Groups.Add(new RichEditExample("Split a Document", string.Empty, string.Empty, BasicActions.SplitDocumentAction, false));
            examples[0].Groups.Add(new RichEditExample("Save a Document", string.Empty, string.Empty, BasicActions.SaveDocumentAction, false));
            examples[0].Groups.Add(new RichEditExample("Print a Document", string.Empty, string.Empty, BasicActions.PrintDocumentAction, false));

            //Add nodes to the "Bookmarks and Hyperlinks" group of examples.
            examples[1].Groups.Add(new RichEditExample("Insert a Bookmark", string.Empty, string.Empty, BookmarksAndHyperlinksActions.InsertBookmarkAction, true));
            examples[1].Groups.Add(new RichEditExample("Insert a Hyperlink", string.Empty, string.Empty, BookmarksAndHyperlinksActions.InsertHyperlinkAction, true));

            //Add nodes to the "Comments" group of examples.
            examples[2].Groups.Add(new RichEditExample("Create a Comment", string.Empty, string.Empty, CommentsActions.CreateCommentAction, true));
            examples[2].Groups.Add(new RichEditExample("Create a Nested Comment", string.Empty, string.Empty, CommentsActions.CreateNestedCommentAction, true));
            examples[2].Groups.Add(new RichEditExample("Delete a Comment", string.Empty, string.Empty, CommentsActions.DeleteCommentAction, true));
            examples[2].Groups.Add(new RichEditExample("Edit Comment Properties", string.Empty, string.Empty, CommentsActions.EditCommentPropertiesAction, true));
            examples[2].Groups.Add(new RichEditExample("Edit Comment Content", string.Empty, string.Empty, CommentsActions.EditCommentContentAction, true));

            //Add nodes to the "Content Controls" group of examples.
            examples[3].Groups.Add(new RichEditExample("Create Content Controls", string.Empty, string.Empty, ContentControlsActions.CreateContentControlsAction, true));
            examples[3].Groups.Add(new RichEditExample("Change Content Control Parameters", string.Empty, string.Empty, ContentControlsActions.ChangeContentControlsAction, true));
            examples[3].Groups.Add(new RichEditExample("Remove Content Controls", string.Empty, string.Empty, ContentControlsActions.RemoveContentControlsAction, true));




            //Add nodes to the "Custom XML parts" group of examples.
            examples[4].Groups.Add(new RichEditExample("Add a Custom Xml Part", string.Empty, string.Empty, CustomXmlActions.AddCustomXmlPartAction, true));
            examples[4].Groups.Add(new RichEditExample("Access a Custom Xml Part", string.Empty, string.Empty, CustomXmlActions.AccessCustomXmlPartAction, true));
            examples[4].Groups.Add(new RichEditExample("Remove a Custom Xml Part", string.Empty, string.Empty, CustomXmlActions.RemoveCustomXmlPartAction, true));

            //Add nodes to the "Document Properties" group of examples.
            examples[5].Groups.Add(new RichEditExample("Set Built-in Properties", string.Empty, string.Empty, DocumentPropertiesActions.StandardDocumentPropertiesAction, true));
            examples[5].Groups.Add(new RichEditExample("Set Custom Properties", string.Empty, string.Empty, DocumentPropertiesActions.CustomDocumentPropertiesAction, true));

            //Add nodes to the "Export" group of examples.
            examples[6].Groups.Add(new RichEditExample("Export a Range to HTML", string.Empty, string.Empty, ExportActions.ExportRangeToHtmlAction, false));
            examples[6].Groups.Add(new RichEditExample("Export a Range to Plain Text", string.Empty, string.Empty, ExportActions.ExportRangeToPlainTextAction, false));
            examples[6].Groups.Add(new RichEditExample("Convert DOCX to PDF", string.Empty, string.Empty, ExportActions.ExportToPDFAction, false));
            examples[6].Groups.Add(new RichEditExample("Convert HTML to PDF", string.Empty, string.Empty, ExportActions.ConvertHTMLtoPDFAction, false));
            examples[6].Groups.Add(new RichEditExample("Convert HTML to DOCX", string.Empty, string.Empty, ExportActions.ConvertHTMLtoDOCXAction, false));
            examples[6].Groups.Add(new RichEditExample("Convert DOCX to HTML", string.Empty, string.Empty, ExportActions.ExportToHTMLAction, false));
            examples[6].Groups.Add(new RichEditExample("Handle the Before Export Event", string.Empty, string.Empty, ExportActions.BeforeExportAction, false));

            //Add nodes to the "Fields" group of examples.
            examples[7].Groups.Add(new RichEditExample("Insert a Field", string.Empty, string.Empty, FieldActions.InsertFieldAction, true));
            examples[7].Groups.Add(new RichEditExample("Modify a Field", string.Empty, string.Empty, FieldActions.ModifyFieldCodeAction, true));
            examples[7].Groups.Add(new RichEditExample("Create a Field from a Range", string.Empty, string.Empty, FieldActions.CreateFieldFromRangeAction, true));

            //Add nodes to the "Formatting" group of examples.
            examples[8].Groups.Add(new RichEditExample("Format Text", string.Empty, string.Empty, FormattingActions.FormatTextAction, true));
            examples[8].Groups.Add(new RichEditExample("Change Spacing", string.Empty, string.Empty, FormattingActions.ChangeSpacingAction, true));
            examples[8].Groups.Add(new RichEditExample("Reset Character Formatting", string.Empty, string.Empty, FormattingActions.ResetCharacterFormattingAction, true));
            examples[8].Groups.Add(new RichEditExample("Format a Paragraph", string.Empty, string.Empty, FormattingActions.FormatParagraphAction, true));
            examples[8].Groups.Add(new RichEditExample("Reset Paragraph Formatting", string.Empty, string.Empty, FormattingActions.ResetParagraphFormattingAction, true));

            //Add nodes to the "Form Fields" group of examples.
            examples[9].Groups.Add(new RichEditExample("Insert a CheckBox", string.Empty, string.Empty, FormFieldsActions.InsertCheckBoxAction, true));

            //Add nodes to the "Headers and Footers" group of examples.
            examples[10].Groups.Add(new RichEditExample("Create a Header", string.Empty, string.Empty, HeadersAndFootersActions.CreateHeaderAction, true));
            examples[10].Groups.Add(new RichEditExample("Modify a Header", string.Empty, string.Empty, HeadersAndFootersActions.ModifyHeaderAction, true));

            //Add nodes to the "Import" group of examples.
            examples[11].Groups.Add(new RichEditExample("Import RTF Text", string.Empty, string.Empty, ImportActions.ImportRtfTextAction, true));
            examples[11].Groups.Add(new RichEditExample("Handle the Before Import Event", string.Empty, string.Empty, ImportActions.BeforeImportAction, true));

            //Add nodes to the "Inline Pictures" group of examples.
            examples[12].Groups.Add(new RichEditExample("Access an Image Collection", string.Empty, string.Empty, InlinePicturesActions.ImageCollectionAction, true));
            examples[12].Groups.Add(new RichEditExample("Save an Image to a File", string.Empty, string.Empty, InlinePicturesActions.SaveImageToFileAction, false));

            //Add nodes to the "Lists" group of examples.
            examples[13].Groups.Add(new RichEditExample("Create a Bulleted List", string.Empty, string.Empty, ListsActions.CreateBulletedListAction, true));
            examples[13].Groups.Add(new RichEditExample("Create a Numbered List", string.Empty, string.Empty, ListsActions.CreateNumberedListAction, true));
            examples[13].Groups.Add(new RichEditExample("Create a Multilevel List", string.Empty, string.Empty, ListsActions.CreateMultilevelListAction, true));

            //Add nodes to the "Notes" group of examples.
            examples[14].Groups.Add(new RichEditExample("Insert Footnotes", string.Empty, string.Empty, NotesActions.InsertFootnotesAction, true));
            examples[14].Groups.Add(new RichEditExample("Insert Endnotes", string.Empty, string.Empty, NotesActions.InsertEndnotesAction, true));
            examples[14].Groups.Add(new RichEditExample("Edit a Footnote", string.Empty, string.Empty, NotesActions.EditFootnoteAction, true));
            examples[14].Groups.Add(new RichEditExample("Edit an Endnote", string.Empty, string.Empty, NotesActions.EditEndnoteAction, true));
            examples[14].Groups.Add(new RichEditExample("Edit a Separator", string.Empty, string.Empty, NotesActions.EditSeparatorAction, true));
            examples[14].Groups.Add(new RichEditExample("Remove Notes", string.Empty, string.Empty, NotesActions.RemoveNotesAction, true));

            //Add nodes to the "Page Layout" group of examples.
            examples[15].Groups.Add(new RichEditExample("Add Line Numbering", string.Empty, string.Empty, PageLayoutActions.LineNumberingAction, true));
            examples[15].Groups.Add(new RichEditExample("Create Columns", string.Empty, string.Empty, PageLayoutActions.CreateColumnsAction, true));
            examples[15].Groups.Add(new RichEditExample("Adjust Page Layout", string.Empty, string.Empty, PageLayoutActions.PrintLayoutAction, true));
            examples[15].Groups.Add(new RichEditExample("Set Tab Stops", string.Empty, string.Empty, PageLayoutActions.TabStopsAction, true));

            //Add nodes to the "Protection" group of examples.
            examples[16].Groups.Add(new RichEditExample("Protect a Document", string.Empty, string.Empty, ProtectionActions.ProtectDocumentAction, false));
            examples[16].Groups.Add(new RichEditExample("Unprotect a Document", string.Empty, string.Empty, ProtectionActions.UnprotectDocumentAction, false));
            examples[16].Groups.Add(new RichEditExample("Create Range Permissions", string.Empty, string.Empty, ProtectionActions.CreateRangePermissionsAction, false));

            //Add nodes to the "Ranges" group of examples.
            examples[17].Groups.Add(new RichEditExample("Insert Text in a Range", string.Empty, string.Empty, RangeActions.InsertTextInRangeAction, true));
            examples[17].Groups.Add(new RichEditExample("Append Text to a Range", string.Empty, string.Empty, RangeActions.AppendTextToRangeAction, true));
            examples[17].Groups.Add(new RichEditExample("Append Text to a Paragraph", string.Empty, string.Empty, RangeActions.AppendToParagraphAction, true));

            //Add nodes to the "Shapes" group of examples.
            examples[18].Groups.Add(new RichEditExample("Add a Floating Picture", string.Empty, string.Empty, ShapesActions.AddFloatingPictureAction, true));
            examples[18].Groups.Add(new RichEditExample("Floating Picture Offset", string.Empty, string.Empty, ShapesActions.FloatingPictureOffsetAction, true));
            examples[18].Groups.Add(new RichEditExample("Change Z-Order and Wrapping", string.Empty, string.Empty, ShapesActions.ChangeZorderAndWrappingAction, true));
            examples[18].Groups.Add(new RichEditExample("Add a Text Box", string.Empty, string.Empty, ShapesActions.AddTextBoxAction, true));
            examples[18].Groups.Add(new RichEditExample("Insert Rich Text in a TextBox", string.Empty, string.Empty, ShapesActions.InsertRichTextInTextBoxAction, true));
            examples[18].Groups.Add(new RichEditExample("Rotate and Resize Shapes", string.Empty, string.Empty, ShapesActions.RotateAndResizeAction, true));

            //Add nodes to the "Styles" group of examples.
            examples[19].Groups.Add(new RichEditExample("Create a New Character Style", string.Empty, string.Empty, StylesAction.CreateNewCharacterStyleAction, true));
            examples[19].Groups.Add(new RichEditExample("Create a New Paragraph Style", string.Empty, string.Empty, StylesAction.CreateNewParagraphStyleAction, true));
            examples[19].Groups.Add(new RichEditExample("Create a New Linked Style", string.Empty, string.Empty, StylesAction.CreateNewLinkedStyleAction, false));

            //Add nodes to the "Tables" group of examples.
            examples[20].Groups.Add(new RichEditExample("Create a Table", string.Empty, string.Empty, TablesActions.CreateTableAction, true));
            examples[20].Groups.Add(new RichEditExample("Create a Fixed Table", string.Empty, string.Empty, TablesActions.CreateFixedTableAction, true));
            examples[20].Groups.Add(new RichEditExample("Change the Table Color", string.Empty, string.Empty, TablesActions.ChangeTableColorAction, true));
            examples[20].Groups.Add(new RichEditExample("Create and Apply a Table Style", string.Empty, string.Empty, TablesActions.CreateAndApplyTableStyleAction, true));
            examples[20].Groups.Add(new RichEditExample("Use a Conditional Style", string.Empty, string.Empty, TablesActions.UseConditionalStyleAction, true));
            examples[20].Groups.Add(new RichEditExample("Change Column Appearance", string.Empty, string.Empty, TablesActions.ChangeColumnAppearanceAction, true));
            examples[20].Groups.Add(new RichEditExample("Table Cell Processor", string.Empty, string.Empty, TablesActions.UseTableCellProcessorAction, true));
            examples[20].Groups.Add(new RichEditExample("Merge Cells", string.Empty, string.Empty, TablesActions.MergeCellsAction, true));
            examples[20].Groups.Add(new RichEditExample("Split Cells", string.Empty, string.Empty, TablesActions.SplitCellsAction, true));
            examples[20].Groups.Add(new RichEditExample("Delete Table Elements", string.Empty, string.Empty, TablesActions.DeleteTableElementsAction, true));
            examples[20].Groups.Add(new RichEditExample("Wrap Text Around a Table", string.Empty, string.Empty, TablesActions.WrapTextAroundTableAction, true));

            //Add nodes to the "Watermarks" group of examples.
            examples[21].Groups.Add(new RichEditExample("Create a Text Watermark", string.Empty, string.Empty, WatermarkActions.CreateTextWatermarkAction, true));
            examples[21].Groups.Add(new RichEditExample("Create an Image Watermark", string.Empty, string.Empty, WatermarkActions.CreateImageWatermarkAction, true));

            return examples;
            #endregion
        }
        public static Dictionary<string, FileInfo> GatherExamplesFromProject(string examplesPath, ExampleLanguage language)
        {
            Dictionary<string, FileInfo> result = new Dictionary<string, FileInfo>();
            foreach (string fileName in Directory.GetFiles(examplesPath, "*" + GetCodeExampleFileExtension(language)))
                result.Add(Path.GetFileNameWithoutExtension(fileName), new FileInfo(fileName));
            return result;
        }
        public static string GetCodeExampleFileExtension(ExampleLanguage language)
        {
            if (language == ExampleLanguage.VB)
                return ".vb";
            return ".cs";
        }
        public static string[] DeleteLeadingWhiteSpaces(string[] lines, String stringToDelete)
        {
            string[] result = new string[lines.Length];
            int stringToDeleteLength = stringToDelete.Length;

            for (int i = 0; i < lines.Length; i++)
            {
                int index = lines[i].IndexOf(stringToDelete);
                result[i] = (index >= 0) ? lines[i].Substring(index + stringToDeleteLength) : lines[i];
            }
            return result;
        }
        public static string ConvertStringToHumanReadableForm(string exampleName)
        {
            string result = SplitCamelCase(exampleName);
            result = result.Replace(" In ", " in ");
            result = result.Replace(" And ", " and ");
            result = result.Replace(" To ", " to ");
            result = result.Replace(" From ", " from ");
            result = result.Replace(" With ", " with ");
            result = result.Replace(" By ", " by ");
            result = result.Replace('\"', '\0');
            return result;
        }
        static string SplitCamelCase(string exampleName)
        {
            int length = exampleName.Length;
            if (length == 1)
                return exampleName;

            StringBuilder result = new StringBuilder(length * 2);
            for (int position = 0; position < length - 1; position++)
            {
                char current = exampleName[position];
                char next = exampleName[position + 1];
                result.Append(current);
                if (char.IsLower(current) && char.IsUpper(next))
                {
                    result.Append(' ');
                }
            }
            result.Append(exampleName[length - 1]);
            return result.ToString();
        }
        public static string GetExamplePath(string exampleFolderName)
        {
            string examplesPath2 = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\", exampleFolderName);
            if (Directory.Exists(examplesPath2))
                return examplesPath2;
            string examplesPathInInsallation = GetRelativeDirectoryPath(exampleFolderName);
            return examplesPathInInsallation;
        }

        public static string GetRelativeDirectoryPath(string name)
        {
            name = "Data\\" + name;
            string path = System.Windows.Forms.Application.StartupPath;
            string s = "\\";
            for (int i = 0; i <= 10; i++)
            {
                if (System.IO.Directory.Exists(path + s + name))
                    return (path + s + name);
                else
                    s += "..\\";
            }
            return "";
        }
        public static GroupsOfRichEditExamples FindExamples(Dictionary<string, FileInfo> examples, ExampleFinder exampleFinder)
        {
            GroupsOfRichEditExamples richEditExamples = InitData();
            foreach (KeyValuePair<string, FileInfo> sourceCodeItem in examples)
            {
                string key = sourceCodeItem.Key;
                List<CodeExample> foundExamples = exampleFinder.Process(examples[key]);
                if (foundExamples.Count == 0)
                    continue;
                foreach (RichEditNode node in richEditExamples)
                {
                    if (node.Name == foundExamples[0].HumanReadableGroupName && node.Groups.Count == foundExamples.Count)
                    {
                        int i = 0;
                        foreach (RichEditExample example in node.Groups)
                        {
                            example.CodeCS = foundExamples[i].CodeCS;
                            example.CodeVB = foundExamples[i].CodeVB;
                            i++;
                        }
                    }
                }
            }
            return richEditExamples;
        }

        public static ExampleLanguage DetectExampleLanguage(string solutionFileNameWithoutExtenstion)
        {
            string projectPath = Directory.GetCurrentDirectory() + "\\..\\..\\";

            string[] csproject = Directory.GetFiles(projectPath, "*.csproj");
            if (csproject.Length != 0 && csproject[0].EndsWith(solutionFileNameWithoutExtenstion + ".csproj"))
                return ExampleLanguage.Csharp;
            string[] vbproject = Directory.GetFiles(projectPath, "*.vbproj");
            if (vbproject.Length != 0 && vbproject[0].EndsWith(solutionFileNameWithoutExtenstion + ".vbproj"))
                return ExampleLanguage.VB;
            return ExampleLanguage.Csharp;
        }
    }
    #endregion

}

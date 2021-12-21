using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using DevExpress.XtraRichEdit;
using DevExpress.XtraTab;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using RichEditDocumentServerAPIExample.CodeUtils;

namespace RichEditDocumentServerAPIExample
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        ExampleCodeEditor codeEditor;
        RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
        GroupsOfRichEditExamples richEditExamples = new GroupsOfRichEditExamples();

        public Form1()
        {
            InitializeComponent();
            InitExamples();
            ShowExamplesInTreeList(treeList1);
            this.codeEditor = new ExampleCodeEditor(richEditControlCS, richEditControlVB);
            InitCurrentExampleLanguage();
            InitTreeListControl(richEditExamples);
            ShowFirstExample();
        }
        void InitExamples()
        {
            string examplePath = CodeExampleUtils.GetExamplePath("CodeExamples");

            Dictionary<string, FileInfo> examplesCS = CodeExampleUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp);
            Dictionary<string, FileInfo> examplesVB = CodeExampleUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.VB);
            DisableTabs(examplesCS.Count, examplesVB.Count);
            Dictionary<string, FileInfo> actualExamples;
            ExampleFinder exampleFinder;
            if (examplesCS.Count != 0)
            {
                actualExamples = examplesCS;
                exampleFinder = new ExampleFinderCSharp();
            }
            else
            {
                actualExamples = examplesVB;
                exampleFinder = new ExampleFinderVB();
            }
            this.richEditExamples = CodeExampleUtils.FindExamples(actualExamples, exampleFinder);
        }
        void InitCurrentExampleLanguage()
        {
            ExampleLanguage currentLanguage = CodeExampleUtils.DetectExampleLanguage("RichEditDocumentServerAPIExample");
            this.codeEditor.CurrentExampleLanguage = currentLanguage;
            xtraTabControl1.SelectedTabPageIndex = (currentLanguage == ExampleLanguage.Csharp) ? 0 : 1;
        }
        void InitTreeListControl(GroupsOfRichEditExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
        }        

        void ShowExamplesInTreeList(TreeList treeList)
        {
            #region InitializeTreeList
            treeList.OptionsPrint.UsePrintStyles = true;
            treeList.FocusedNodeChanged += new FocusedNodeChangedEventHandler(this.OnNewExampleSelected);
            treeList.OptionsView.ShowColumns = false;
            treeList.OptionsView.ShowIndicator = false;
            #endregion
            TreeListColumn col1 = new TreeListColumn();
            col1.Caption = "Name";
            col1.VisibleIndex = 0;
            col1.OptionsColumn.AllowEdit = false;
            col1.OptionsColumn.AllowMove = false;
            col1.OptionsColumn.ReadOnly = true;
            treeList.Columns.AddRange(new TreeListColumn[] { col1 });
        }

        void ShowFirstExample()
        {
            treeList1.ExpandAll();
            if (treeList1.Nodes.Count > 0)
                treeList1.FocusedNode = treeList1.MoveFirst().FirstNode;
            RichEditExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as RichEditExample;
            codeEditor.ShowExample(example);

        }

        void OnNewExampleSelected(object sender, FocusedNodeChangedEventArgs e)
        {
            RichEditExample codeExample = (sender as TreeList).GetDataRecordByNode(e.Node) as RichEditExample;

            if (codeExample == null)
                return;

            codeEditor.ShowExample(codeExample);
            codeExampleNameLbl.Text = CodeExampleUtils.ConvertStringToHumanReadableForm(codeExample.Name);
        }

        void DisableTabs(int examplesCSCount, int examplesVBCount)
        {
            if (examplesCSCount == 0)
                foreach (XtraTabPage t in xtraTabControl1.TabPages) if (t.Tag.ToString() == "CS") t.PageEnabled = false;
            if (examplesVBCount == 0)
                foreach (XtraTabPage t in xtraTabControl1.TabPages) if (t.Tag.ToString() == "VB") t.PageEnabled = false;
        }

        void OnRunButtonClick(object sender, EventArgs e)
        {
            wordProcessor.CreateNewDocument();
            RichEditExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as RichEditExample;
            if (example == null)
                return;
            Action<RichEditDocumentServer> action = example.Action;
            action(wordProcessor);
            SaveDocumentToFile(example);
        }
        void SaveDocumentToFile(RichEditExample example)
        {            
            // Save the modified document to the file.
            if (example.SaveResult)
            {
                try
                {
                    wordProcessor.SaveDocument("Result.docx", DocumentFormat.OpenXml);
                    Process.Start("Result.docx");
                }
                catch (Exception)
                {
                    MessageBox.Show("Close the Result.docx file.");
                }
            }
        }
    }   
}

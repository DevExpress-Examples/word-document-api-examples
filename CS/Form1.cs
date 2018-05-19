using DevExpress.XtraTab;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using RichEditDocumentServerAPIExample.CodeUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample
{
    public partial class Form1 : Form
    {
        #region InitializeComponent
        #endregion

        ExampleCodeEditor codeEditor;
        ExampleEvaluatorByTimer evaluator;
        List<CodeExampleGroup> examples;
        bool treeListRootNodeLoading = true;
        RichEditDocumentServer server = new RichEditDocumentServer();

        public Form1()
        {
            InitializeComponent();

            string examplePath = CodeExampleDemoUtils.GetExamplePath("CodeExamples");

            Dictionary<string, FileInfo> examplesCS = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.Csharp);
            Dictionary<string, FileInfo> examplesVB = CodeExampleDemoUtils.GatherExamplesFromProject(examplePath, ExampleLanguage.VB);
            DisableTabs(examplesCS.Count, examplesVB.Count);
            this.examples = CodeExampleDemoUtils.FindExamples(examplePath, examplesCS, examplesVB);
            ShowExamplesInTreeList(treeList1, examples);

            this.codeEditor = new ExampleCodeEditor(richEditControlCS, richEditControlVB, richEditControlCSClass, richEditControlVBClass);
            CurrentExampleLanguage = CodeExampleDemoUtils.DetectExampleLanguage("RichEditDocumentServerExample");
            this.evaluator = new RichEditExampleEvaluatorByTimer();

            this.evaluator.QueryEvaluate += OnExampleEvaluatorQueryEvaluate;
            this.evaluator.OnBeforeCompile += evaluator_OnBeforeCompile;
            this.evaluator.OnAfterCompile += evaluator_OnAfterCompile;
            this.xtraTabControl1.SelectedPageChanged += xtraTabControl1_SelectedPageChanged;

            ShowFirstExample();
            treeList1.ExpandAll();
        }



        ExampleLanguage CurrentExampleLanguage
        {
            get
            {
                if (xtraTabControl1.SelectedTabPage.Tag.ToString() == "CS") return ExampleLanguage.Csharp;
                else return ExampleLanguage.VB;
            }
            set
            {
                this.codeEditor.CurrentExampleLanguage = value;
                //xtraTabControl1.SelectedTabPageIndex = (value == ExampleLanguage.Csharp) ? 0 : 1;
            }
        }

        void ShowExamplesInTreeList(TreeList treeList, List<CodeExampleGroup> examples)
        {
            #region InitializeTreeList
            treeList.OptionsPrint.UsePrintStyles = true;
            treeList.FocusedNodeChanged += new FocusedNodeChangedEventHandler(this.OnNewExampleSelected);
            treeList.OptionsView.ShowColumns = false;
            treeList.OptionsView.ShowIndicator = false;
            treeList.VirtualTreeGetChildNodes += treeList_VirtualTreeGetChildNodes;
            treeList.VirtualTreeGetCellValue += treeList_VirtualTreeGetCellValue;
            #endregion
            TreeListColumn col1 = new TreeListColumn();
            col1.VisibleIndex = 0;
            col1.OptionsColumn.AllowEdit = false;
            col1.OptionsColumn.AllowMove = false;
            col1.OptionsColumn.ReadOnly = true;
            treeList.Columns.AddRange(new TreeListColumn[] { col1 });
            treeList.DataSource = new Object();
            treeList.ExpandAll();
        }

        void treeList_VirtualTreeGetCellValue(object sender, VirtualTreeGetCellValueInfo args)
        {
            CodeExampleGroup group = args.Node as CodeExampleGroup;
            if (group != null)
                args.CellData = group.Name;

            CodeExample example = args.Node as CodeExample;
            if (example != null)
                args.CellData = example.RegionName;
        }

        void treeList_VirtualTreeGetChildNodes(object sender, VirtualTreeGetChildNodesInfo args)
        {
            if (treeListRootNodeLoading)
            {
                args.Children = examples;
                treeListRootNodeLoading = false;
            }
            else
            {
                if (args.Node == null)
                    return;
                CodeExampleGroup group = args.Node as CodeExampleGroup;
                if (group != null)
                    args.Children = group.Examples;
            }
        }

        void ShowFirstExample()
        {
            treeList1.ExpandAll();
            if (treeList1.Nodes.Count > 0)
                treeList1.FocusedNode = treeList1.MoveFirst().FirstNode;
        }

        void evaluator_OnAfterCompile(object sender, OnAfterCompileEventArgs args)
        {
            codeEditor.AfterCompile(args.Result);
        }

        void evaluator_OnBeforeCompile(object sender, EventArgs e)
        {
            Document document = server.Document;
            document.BeginUpdate();
            codeEditor.BeforeCompile();
            server.CreateNewDocument();
            document.Unit = DevExpress.Office.DocumentUnit.Document;

        }

        void OnNewExampleSelected(object sender, FocusedNodeChangedEventArgs e)
        {
            CodeExample newExample = (sender as TreeList).GetDataRecordByNode(e.Node) as CodeExample;
            CodeExample oldExample = (sender as TreeList).GetDataRecordByNode(e.OldNode) as CodeExample;

            if (newExample == null)
                return;

            string exampleCode = codeEditor.ShowExample(oldExample, newExample);
            codeExampleNameLbl.Text = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(newExample.RegionName) + " example";
            CodeEvaluationEventArgs args = new CodeEvaluationEventArgs();
            InitializeCodeEvaluationEventArgs(args);
            //evaluator.ForceCompile(args);

        }

        void InitializeCodeEvaluationEventArgs(CodeEvaluationEventArgs e)
        {
            e.Result = true;
            e.Code = codeEditor.CurrentCodeEditor.Text;
            e.CodeClasses = codeEditor.CurrentCodeClassEditor.Text;
            e.Language = CurrentExampleLanguage;
            e.EvaluationParameter = server;
        }

        void OnExampleEvaluatorQueryEvaluate(object sender, CodeEvaluationEventArgs e)
        {
            e.Result = false;
            if (codeEditor.RichEditTextChanged)
            {// && compileComplete) {
                TimeSpan span = DateTime.Now - codeEditor.LastExampleCodeModifiedTime;

                if (span < TimeSpan.FromMilliseconds(1000))
                {//CompileTimeIntervalInMilliseconds  1900
                    codeEditor.ResetLastExampleModifiedTime();
                    return;
                }
                //e.Result = true;
                InitializeCodeEvaluationEventArgs(e);
            }
        }

        void DisableTabs(int examplesCSCount, int examplesVBCount)
        {
            if (examplesCSCount == 0)
                foreach (XtraTabPage t in xtraTabControl1.TabPages) if (t.Tag.ToString() == "CS") t.PageEnabled = false;
            if (examplesVBCount == 0)
                foreach (XtraTabPage t in xtraTabControl1.TabPages) if (t.Tag.ToString() == "VB") t.PageEnabled = false;
        }



        void xtraTabControl1_SelectedPageChanged(object sender, TabPageChangedEventArgs e)
        {
            CurrentExampleLanguage = (e.Page.Tag.ToString() == "CS") ? ExampleLanguage.Csharp : ExampleLanguage.VB;

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            server.SaveDocument("Result.docx", DocumentFormat.OpenXml);
            Process.Start("Result.docx");

        }
    }
}

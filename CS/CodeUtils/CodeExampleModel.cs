using System;
using System.Collections.Generic;
using System.ComponentModel;
using DevExpress.XtraRichEdit;
using DevExpress.XtraTreeList;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    public class RichEditExample : RichEditNode
    {
        public RichEditExample(string name, string codeCS, string codeVB, Action<RichEditDocumentServer> action, bool saveResult) : base(name)
        {
            Action = action;
            CodeCS = codeCS;
            CodeVB = codeVB;
            SaveResult = saveResult;
        }
        public Action<RichEditDocumentServer> Action { get; private set; }
        public string CodeCS { get; set; }
        public string CodeVB { get; set; }
        public bool SaveResult { get; set; }

    }
    public class RichEditNode
    {
        GroupsOfRichEditExamples groups = new GroupsOfRichEditExamples();

        public RichEditNode(string name)
        {
            Name = name;
        }
        [Browsable(false)]
        public GroupsOfRichEditExamples Groups { get { return groups; } }
        public string Name { get; set; }
        [Browsable(false)]
        public GroupsOfRichEditExamples Owner { get; set; }
    }

    public class GroupsOfRichEditExamples : BindingList<RichEditNode>, TreeList.IVirtualTreeListData
    {
        void TreeList.IVirtualTreeListData.VirtualTreeGetChildNodes(VirtualTreeGetChildNodesInfo info)
        {
            RichEditNode obj = info.Node as RichEditNode;
            info.Children = obj.Groups;
        }
        protected override void InsertItem(int index, RichEditNode item)
        {
            item.Owner = this;
            base.InsertItem(index, item);
        }
        void TreeList.IVirtualTreeListData.VirtualTreeGetCellValue(VirtualTreeGetCellValueInfo info)
        {
            RichEditNode obj = info.Node as RichEditNode;
            switch (info.Column.Caption)
            {
                case "Name":
                    info.CellData = obj.Name;
                    break;
            }
        }
        void TreeList.IVirtualTreeListData.VirtualTreeSetCellValue(VirtualTreeSetCellValueInfo info)
        { }
    }

    public class CodeExampleGroup
    {
        public string Name { get; set; }
        public List<CodeExample> Examples { get; set; }
        public int Id { get; set; }
    }

    public class CodeExample
    {
        public string CodeCS { get; set; }
        //public string CodeCsHelper { get; set; }
        public string CodeVB { get; set; }
        //public string CodeVbHelper { get; set; }
        public string RegionName { get; set; }
        public string HumanReadableGroupName { get; set; }
        public string ExampleGroup { get; set; }
        public int Id { get; set; }
    }

    public enum ExampleLanguage
    {
        Csharp = 0,
        VB = 1
    }
}

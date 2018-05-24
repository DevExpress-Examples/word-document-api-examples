using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    public class CodeExampleGroup
    {
        public CodeExampleGroup()
        {
        }
        public string Name { get; set; }
        public List<CodeExample> Examples { get; set; }
        public int Id { get; set; }
    }

    public class CodeExample
    {
        public string CodeCS { get; set; }
        public string CodeCsHelper { get; set; }
        public string CodeVB { get; set; }
        public string CodeVbHelper { get; set; }
        public string RegionName { get; set; }
        public string HumanReadableGroupName { get; set; }
        public string ExampleGroup { get; set; }
        public int Id { get; set; }
    }

    public class CodeExampleCollection : List<CodeExample>
    {
        public void Merge(CodeExample example)
        {
            CodeExample item = this.Find(x => x.HumanReadableGroupName.Equals(example.HumanReadableGroupName)
                    && x.RegionName.Equals(example.RegionName));
            if (item == null)
            {
                item = new CodeExample();
                item.HumanReadableGroupName = example.HumanReadableGroupName;
                item.RegionName = example.RegionName;
                this.Add(item);
            }
            item.CodeCS += example.CodeCS;
            item.CodeCsHelper += example.CodeCsHelper;
            item.CodeVB += example.CodeVB;
            item.CodeVbHelper += example.CodeVbHelper;
        }

        public void Merge(List<CodeExample> exampleList)
        {
            foreach (CodeExample item in exampleList) this.Merge(item);
        }
    }


    public enum ExampleLanguage
    {
        Csharp = 0,
        VB = 1
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    #region ExampleFinder
    public abstract class ExampleFinder
    {
        public bool isHelper = false;
        public abstract string RegexRegionPattern { get; }
        public abstract string RegionStartPattern { get; }
        public abstract string RegionHelperStartPattern { get; }

        public List<CodeExample> Process(FileInfo fileWithExample)
        {
            if (fileWithExample == null)
                return new List<CodeExample>();

            string groupName = Path.GetFileNameWithoutExtension(fileWithExample.Name);
            string code;
            using (FileStream stream = File.Open(fileWithExample.FullName, FileMode.Open, FileAccess.Read))
            {
                StreamReader sr = new StreamReader(stream);
                code = sr.ReadToEnd();
            }
            List<CodeExample> foundExamples = ParseSouceFileAndFindRegionsWithExamples(groupName, code);
            return foundExamples;
        }

        public List<CodeExample> ParseSouceFileAndFindRegionsWithExamples(string groupName, string sourceCode)
        {
            List<CodeExample> result = new List<CodeExample>();

            var matches = Regex.Matches(sourceCode, RegexRegionPattern, RegexOptions.Singleline);

            foreach (var match in matches)
            {
                string[] lines = match.ToString().Split(new string[] { "\n" }, StringSplitOptions.None);

                if (lines.Length <= 2)
                    continue;
                lines = DeleteLeadingWhiteSpacesFromSourceCode(lines);

                string regionName = String.Empty;
                bool regionIsValid = ValidateRegionName(lines, ref regionName);
                if (!regionIsValid)
                    continue;

                string exampleCode = string.Join("\r\n", lines, 1, lines.Length - 2);
                result.Add(CreateRichEditExample(groupName, regionName, exampleCode));

            }
            return result;
        }

        protected CodeExample CreateRichEditExample(string exampleGroup, string regionName, string exampleCode)
        {
            CodeExample result = new CodeExample();
            SetExampleCode(exampleCode, result);
            result.RegionName = regionName;
            result.HumanReadableGroupName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(exampleGroup);
            return result;
        }
        protected abstract void SetExampleCode(string exampleCode, CodeExample newExample);

        protected virtual string[] DeleteLeadingWhiteSpacesFromSourceCode(string[] lines)
        {
            return CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, "        ");
        }
        protected virtual bool ValidateRegionName(string[] lines, ref string regionName)
        {
            int keepHashMark = 0; // "#example" if value is -1 or "example" if value will be 0

            string region = lines[0];
            int regionIndex = region.IndexOf(RegionHelperStartPattern);

            if (regionIndex == 0)
            {
                isHelper = true;
                regionName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(region.Substring(regionIndex + RegionHelperStartPattern.Length + keepHashMark));
            }

            if (regionIndex < 0)
            {
                isHelper = false;
                regionIndex = region.IndexOf(RegionStartPattern);
                if (regionIndex < 0)
                {
                    regionName = String.Empty;
                    return false;
                }
                regionName = CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(region.Substring(regionIndex + RegionStartPattern.Length + keepHashMark));
            }
            return true;
        }
    }
    #endregion
    #region ExampleFinderVB
    public class ExampleFinderVB : ExampleFinder
    {
        //public ExampleFinderVB() {
        //}
        public override string RegexRegionPattern { get { return "#Region.*?#End Region"; } }
        public override string RegionStartPattern { get { return "#Region \"#"; } }
        public override string RegionHelperStartPattern { get { return "#Region \"#@"; } }

        protected override string[] DeleteLeadingWhiteSpacesFromSourceCode(string[] lines)
        {
            string[] result = base.DeleteLeadingWhiteSpacesFromSourceCode(lines);
            return CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(result, "\t\t");
        }
        protected override bool ValidateRegionName(string[] lines, ref string regionName)
        {
            bool result = base.ValidateRegionName(lines, ref regionName);
            if (!result)
                return result;
            regionName = regionName.TrimEnd('\"');
            return true;
        }
        protected override void SetExampleCode(string code, CodeExample newExample)
        {
            if (isHelper)
                newExample.CodeVbHelper = code;
            else
                newExample.CodeVB = code;
        }
    }
    #endregion
    #region ExampleFinderCSharp
    public class ExampleFinderCSharp : ExampleFinder
    {
        public override string RegexRegionPattern { get { return "#region.*?#endregion"; } }
        public override string RegionStartPattern { get { return "#region #"; } }
        public override string RegionHelperStartPattern { get { return "#region #@"; } }

        protected override void SetExampleCode(string code, CodeExample newExample)
        {
            if (isHelper)
                newExample.CodeCsHelper = code;
            else
                newExample.CodeCS = code;
        }
    }
    #endregion
}

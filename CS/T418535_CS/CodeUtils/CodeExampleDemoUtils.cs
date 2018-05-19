using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    #region CodeExampleDemoUtils
    public static class CodeExampleDemoUtils
    {
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
        public static string ConvertStringToMoreHumanReadableForm(string exampleName)
        {
            string result = SplitCamelCase(exampleName);
            result = result.Replace(" In ", " in ");
            result = result.Replace(" And ", " and ");
            result = result.Replace(" To ", " to ");
            result = result.Replace(" From ", " from ");
            result = result.Replace(" With ", " with ");
            result = result.Replace(" By ", " by ");
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
        {//"CodeExamples"
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
        public static List<CodeExampleGroup> FindExamples(string examplePath, Dictionary<string, FileInfo> examplesCS, Dictionary<string, FileInfo> examplesVB)
        {

            List<CodeExampleGroup> result = new List<CodeExampleGroup>();

            Dictionary<string, FileInfo> current = null;
            ExampleFinder csExampleFinder = new ExampleFinderCSharp();
            ExampleFinder vbExampleFinder = new ExampleFinderVB();

            current = (examplesCS.Count != 0) ? examplesCS : examplesVB;

            foreach (KeyValuePair<string, FileInfo> sourceCodeItem in current)
            {
                string key = sourceCodeItem.Key;

                List<CodeExample> foundExamplesCS = new List<CodeExample>();
                if (examplesCS.Count != 0)
                    foundExamplesCS = csExampleFinder.Process(examplesCS[key]);

                List<CodeExample> foundExamplesVB = new List<CodeExample>();
                if (examplesVB.Count != 0)
                    foundExamplesVB = vbExampleFinder.Process(examplesVB[key]);

                CodeExampleCollection mergedExamplesCollection = new CodeExampleCollection();

                mergedExamplesCollection.Merge(foundExamplesCS);
                mergedExamplesCollection.Merge(foundExamplesVB);

                if (mergedExamplesCollection.Count == 0)
                    continue;

                CodeExampleGroup group = new CodeExampleGroup()
                {
                    Name = mergedExamplesCollection[0].HumanReadableGroupName,
                    Examples = mergedExamplesCollection
                };
                result.Add(group);
            }
            return result;
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

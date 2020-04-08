using System;
using System.CodeDom.Compiler;
using System.Reflection;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    public abstract class ExampleCodeEvaluator
    {
        protected abstract string CodeStart { get; }
        protected abstract string CodeBeforeClasses { get; }
        protected abstract string CodeEnd { get; }
        protected abstract CodeDomProvider GetCodeDomProvider();
        protected abstract string GetModuleAssembly();
        protected abstract string GetExampleClassName();

        public bool ExecuteCodeAndGenerateDocument(CodeEvaluationEventArgs args)
        {
            string theCode = String.Concat(CodeStart, args.Code, CodeBeforeClasses, args.CodeClasses, CodeEnd);
            string[] linesOfCode = new string[] { theCode };
            return CompileAndRun(linesOfCode, args.EvaluationParameter);
        }

        protected internal bool CompileAndRun(string[] linesOfCode, object evaluationParameter)
        {
            CompilerParameters CompilerParams = new CompilerParameters();

            CompilerParams.GenerateInMemory = true;
            CompilerParams.TreatWarningsAsErrors = false;
            CompilerParams.GenerateExecutable = false;

            string[] referencesSystem = new string[] { "System.dll",
                                                      "System.Windows.Forms.dll",
                                                      "System.Data.dll",
                                                      "System.Xml.dll",
                                                      "System.Core.dll",
                                                      "System.Drawing.dll" };

            string[] referencesDX = new string[] {
                AssemblyInfo.SRAssemblyData,
                GetModuleAssembly(),
                AssemblyInfo.SRAssemblyOfficeCore,
                AssemblyInfo.SRAssemblyPrintingCore,
                AssemblyInfo.SRAssemblyPrinting,
                AssemblyInfo.SRAssemblyDocs,
                AssemblyInfo.SRAssemblyUtils,
                AssemblyInfo.SRAssemblyRichEdit
            };
            string[] references = new string[referencesSystem.Length + referencesDX.Length];

            for (int referenceIndex = 0; referenceIndex < referencesSystem.Length; referenceIndex++)
            {
                references[referenceIndex] = referencesSystem[referenceIndex];
            }

            for (int i = 0, initial = referencesSystem.Length; i < referencesDX.Length; i++)
            {
                Assembly assembly = Assembly.Load(referencesDX[i] + AssemblyInfo.FullAssemblyVersionExtension);
                if (assembly != null)
                    references[i + initial] = assembly.Location;
            }
            CompilerParams.ReferencedAssemblies.AddRange(references);


            CodeDomProvider provider = GetCodeDomProvider();
            CompilerResults compile = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode);

            if (compile.Errors.HasErrors)
            {
                string text = "Compile error: ";
                foreach (CompilerError ce in compile.Errors)
                {
                    text += "rn" + ce.ToString();
                }
                System.Windows.Forms.MessageBox.Show(text);
                return false;
            }

            Module module = null;
            try
            {
                module = compile.CompiledAssembly.GetModules()[0];
            }
            catch
            {
            }
            Type moduleType = null;
            if (module == null)
            {
                return false;
            }
            moduleType = module.GetType(GetExampleClassName());

            MethodInfo methInfo = null;
            if (moduleType == null)
            {
                return false;
            }
            methInfo = moduleType.GetMethod("Process");

            if (methInfo != null)
            {
                try
                {
                    methInfo.Invoke(null, new object[] { evaluationParameter });
                }
                catch (Exception)
                {
                    return false;// an error
                }
                return true;
            }
            return false;
        }
    }

    public abstract class RichEditExampleCodeEvaluator : ExampleCodeEvaluator
    {


        protected override string GetModuleAssembly()
        {
            return AssemblyInfo.SRAssemblyRichEditCore;
        }
        protected override string GetExampleClassName()
        {
            return "RichEditCodeResultViewer.ExampleItem";
        }
    }
    #region RichEditCSExampleCodeEvaluator
    public class RichEditCSExampleCodeEvaluator : RichEditExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new Microsoft.CSharp.CSharpCodeProvider();
        }
        const string codeStart =
      "using System;\r\n" +
            "using System.Text;\r\n" +
      "using DevExpress.XtraPrinting;\r\n" +
      "using DevExpress.XtraRichEdit;\r\n" +
      "using DevExpress.XtraRichEdit.API.Native;\r\n" +
            "using DevExpress.XtraRichEdit.Export;\r\n" +
            "using DevExpress.XtraRichEdit.Import;\r\n" +
            "using DevExpress.Office.Utils;\r\n" +
      "using System.Drawing;\r\n" +
      "using System.Windows.Forms;\r\n" +
      "using DevExpress.Utils;\r\n" +
      "using System.IO;\r\n" +
      "using System.Diagnostics;\r\n" +
      "using System.Xml;\r\n" +
      "using System.Data;\r\n" +
      "using System.Collections.Generic;\r\n" +
      "using System.Linq;\r\n" +
      "using System.Globalization;\r\n" +
      "using Document = DevExpress.XtraRichEdit.API.Native.Document;\r\n" +
      "using TableRow = DevExpress.XtraRichEdit.API.Native.TableRow;\r\n" +
      "namespace RichEditCodeResultViewer { \r\n" +
      "public class ExampleItem { \r\n" +
      "        public static void Process(RichEditDocumentServer wordProcessor) { \r\n" +
      "\r\n";

        const string codeBeforeClasses =
            "       \r\n }\r\n" +
            "    }\r\n";

        const string codeEnd =
            "\r\n" +
            "    }\r\n";

        protected override string CodeStart { get { return codeStart; } }
        protected override string CodeBeforeClasses
        {
            get { return codeBeforeClasses; }
        }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion
    #region RichEditVbExampleCodeEvaluator
    public class RichEditVbExampleCodeEvaluator : RichEditExampleCodeEvaluator
    {

        protected override CodeDomProvider GetCodeDomProvider()
        {
            return new Microsoft.VisualBasic.VBCodeProvider();
        }
        const string codeStart =
      "Imports Microsoft.VisualBasic\r\n" +
      "Imports System\r\n" +
      "Imports DevExpress.XtraRichEdit\r\n" +
      "Imports DevExpress.XtraRichEdit.API.Native\r\n" +
      "Imports System.Drawing\r\n" +
      "Imports System.Windows.Forms\r\n" +
      "Imports DevExpress.Utils\r\n" +
      "Imports System.IO\r\n" +
      "Imports System.Diagnostics\r\n" +
      "Imports System.Xml\r\n" +
      "Imports System.Data\r\n" +
      "Imports System.Collections.Generic\r\n" +
      "Imports System.Globalization\r\n" +
      "Imports Document = DevExpress.XtraRichEdit.API.Native.Document\r\n"+
      "Imports TableRow = DevExpress.XtraRichEdit.API.Native.TableRow\r\n"+
      "Namespace RichEditCodeResultViewer\r\n" +
      "	Public Class ExampleItem\r\n" +
      "		Public Shared Sub Process(ByVal wordProcessor As RichEditDocumentServer)\r\n" +
      "\r\n";

        const string codeBeforeClasses =
            "\r\n		End Sub\r\n" +
            "	End Class\r\n";

        const string codeEnd =
        "\r\nEnd Namespace\r\n";

        protected override string CodeStart { get { return codeStart; } }
        protected override string CodeBeforeClasses
        {
            get { return codeBeforeClasses; }
        }
        protected override string CodeEnd { get { return codeEnd; } }
    }
    #endregion

    public abstract class ExampleEvaluatorByTimer : IDisposable
    {
        LeakSafeCompileEventRouter leakSafeCompileEventRouter;
        System.Windows.Forms.Timer compileExampleTimer;
        bool compileComplete = true;
        const int CompileTimeIntervalInMilliseconds = 2000;

        public ExampleEvaluatorByTimer(bool enableTimer)
        {
            this.leakSafeCompileEventRouter = new LeakSafeCompileEventRouter(this);

            if (enableTimer)
            {
                this.compileExampleTimer = new System.Windows.Forms.Timer();
                this.compileExampleTimer.Interval = CompileTimeIntervalInMilliseconds;

                this.compileExampleTimer.Tick += new EventHandler(leakSafeCompileEventRouter.OnCompileExampleTimerTick); //OnCompileTimerTick
                this.compileExampleTimer.Enabled = true;
            }
        }
        public ExampleEvaluatorByTimer()
            : this(true)
        {
        }

        #region Events
        public event CodeEvaluationEventHandler QueryEvaluate;

        protected internal virtual CodeEvaluationEventArgs RaiseQueryEvaluate()
        {
            if (QueryEvaluate != null)
            {
                CodeEvaluationEventArgs args = new CodeEvaluationEventArgs();
                QueryEvaluate(this, args);
                return args;
            }
            return null;
        }
        public event EventHandler OnBeforeCompile;

        void RaiseOnBeforeCompile()
        {
            if (OnBeforeCompile != null)
                OnBeforeCompile(this, new EventArgs());
        }

        public event OnAfterCompileEventHandler OnAfterCompile;

        void RaiseOnAfterCompile(bool result)
        {
            if (OnAfterCompile != null)
                OnAfterCompile(this, new OnAfterCompileEventArgs() { Result = result });
        }
        #endregion

        public void CompileExample(object sender, EventArgs e)
        {
            if (!compileComplete)
                return;
            CodeEvaluationEventArgs args = RaiseQueryEvaluate();
            if (!args.Result)
                return;

            ForceCompile(args);
        }
        public void ForceCompile(CodeEvaluationEventArgs args)
        {
            compileComplete = false;
            if (!String.IsNullOrEmpty(args.Code))
                CompileExampleAndShowPrintPreview(args);

            compileComplete = true;
        }
        void CompileExampleAndShowPrintPreview(CodeEvaluationEventArgs args)
        {
            bool evaluationSucceed = false;
            try
            {
                RaiseOnBeforeCompile();

                evaluationSucceed = Evaluate(args);
            }
            finally
            {
                RaiseOnAfterCompile(evaluationSucceed);
            }
        }

        public bool Evaluate(CodeEvaluationEventArgs args)
        {
            ExampleCodeEvaluator richeditExampleCodeEvaluator = GetExampleCodeEvaluator(args.Language);
            return richeditExampleCodeEvaluator.ExecuteCodeAndGenerateDocument(args);
        }

        protected abstract ExampleCodeEvaluator GetExampleCodeEvaluator(ExampleLanguage language);

        public void Dispose()
        {
            if (compileExampleTimer != null)
            {
                compileExampleTimer.Enabled = false;
                if (leakSafeCompileEventRouter != null)
                    compileExampleTimer.Tick -= new EventHandler(leakSafeCompileEventRouter.OnCompileExampleTimerTick); //OnCompileTimerTick
                compileExampleTimer.Dispose();
                compileExampleTimer = null;
            }
        }
    }

    #region RichEditExampleEvaluatorByTimer
    public class RichEditExampleEvaluatorByTimer : ExampleEvaluatorByTimer
    {
        public RichEditExampleEvaluatorByTimer()
            : base()
        {
        }

        protected override ExampleCodeEvaluator GetExampleCodeEvaluator(ExampleLanguage language)
        {
            if (language == ExampleLanguage.VB)
                return new RichEditVbExampleCodeEvaluator();
            return new RichEditCSExampleCodeEvaluator();
        }
    }
    #endregion

    #region LeakSafeCompileEventRouter
    public class LeakSafeCompileEventRouter
    {
        readonly WeakReference weakControlRef;

        public LeakSafeCompileEventRouter(ExampleEvaluatorByTimer module)
        {
            //Guard.ArgumentNotNull(module, "module");
            this.weakControlRef = new WeakReference(module);
        }
        public void OnCompileExampleTimerTick(object sender, EventArgs e)
        {
            ExampleEvaluatorByTimer module = (ExampleEvaluatorByTimer)weakControlRef.Target;
            if (module != null)
                module.CompileExample(sender, e);
        }
    }
    public class CodeEvaluationEventArgs : EventArgs
    {
        public bool Result { get; set; }
        public string Code { get; set; }
        public string CodeClasses { get; set; }
        public ExampleLanguage Language { get; set; }
        public object EvaluationParameter { get; set; }
    }
    public delegate void CodeEvaluationEventHandler(object sender, CodeEvaluationEventArgs e);

    public class OnAfterCompileEventArgs : EventArgs
    {
        public bool Result { get; set; }
    }
    public delegate void OnAfterCompileEventHandler(object sender, OnAfterCompileEventArgs e);
    #endregion

}

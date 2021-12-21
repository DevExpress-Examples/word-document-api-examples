using System;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Internal;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    public class ExampleCodeEditor
    {
        readonly IRichEditControl codeEditorCs;
        readonly IRichEditControl codeEditorVb;

        ExampleLanguage current;

        public ExampleCodeEditor(IRichEditControl codeEditorCs, IRichEditControl codeEditorVb/*, IRichEditControl codeEditorCsClass, IRichEditControl codeEditorVbClass*/)
        {
            this.codeEditorCs = codeEditorCs;
            this.codeEditorVb = codeEditorVb;

            this.codeEditorCs.InnerControl.InitializeDocument += new System.EventHandler(this.InitializeSyntaxHighlightForCs);
            this.codeEditorVb.InnerControl.InitializeDocument += new System.EventHandler(this.InitializeSyntaxHighlightForVb);
        }

        void InitializeSyntaxHighlightForCs(object sender, EventArgs e)
        {
            InitializeSyntaxHighlight(codeEditorCs, ExampleLanguage.Csharp);
        }
        void InitializeSyntaxHighlightForVb(object sender, EventArgs e)
        {
            InitializeSyntaxHighlight(codeEditorVb, ExampleLanguage.VB);
        }
        void InitializeSyntaxHighlight(IRichEditControl codeEditor, ExampleLanguage language)
        {
            SyntaxHightlightInitializeHelper syntaxHightlightInitializator = new SyntaxHightlightInitializeHelper();
            syntaxHightlightInitializator.Initialize(codeEditor, CodeExampleUtils.GetCodeExampleFileExtension(language));
            DisableRichEditFeatures(codeEditor);
        }

        public InnerRichEditControl CurrentCodeEditor
        {
            get
            {
                if (CurrentExampleLanguage == ExampleLanguage.Csharp)
                    return codeEditorCs.InnerControl;
                else
                    return codeEditorVb.InnerControl;
            }
        }

        public ExampleLanguage CurrentExampleLanguage
        {
            get { return current; }
            set { current = value; }
        }      

        public void ShowExample(RichEditExample codeExample)
        {
            InnerRichEditControl richEditControlCs = codeEditorCs.InnerControl;
            InnerRichEditControl richEditControlVb = codeEditorVb.InnerControl;

            if (codeExample != null)
            {
                richEditControlCs.Text = codeExample.CodeCS;
                richEditControlVb.Text = codeExample.CodeVB;
            }
        }
        void DisableRichEditFeatures(IRichEditControl codeEditor)
        {
            RichEditControlOptionsBase options = codeEditor.InnerDocumentServer.Options;
            options.DocumentCapabilities.Hyperlinks = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.Bulleted = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.Simple = DocumentCapability.Disabled;
            options.DocumentCapabilities.Numbering.MultiLevel = DocumentCapability.Disabled;

            options.DocumentCapabilities.Tables = DocumentCapability.Disabled;
            options.DocumentCapabilities.Bookmarks = DocumentCapability.Disabled;

            options.DocumentCapabilities.CharacterStyle = DocumentCapability.Disabled;
            options.DocumentCapabilities.ParagraphStyle = DocumentCapability.Disabled;
        }
    }
}

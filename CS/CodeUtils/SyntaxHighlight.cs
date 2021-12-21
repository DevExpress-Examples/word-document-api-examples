using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DevExpress.CodeParser;
using DevExpress.Office.Internal;
using DevExpress.Office.Utils;
using DevExpress.Utils;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraRichEdit.Export;
using DevExpress.XtraRichEdit.Import;
using DevExpress.XtraRichEdit.Internal;
using DevExpress.XtraRichEdit.Services;

namespace RichEditDocumentServerAPIExample.CodeUtils
{
    public class SyntaxHightlightInitializeHelper
    {
        public void Initialize(IRichEditControl richEditControl, string codeExamplesFileExtension)
        {
            InnerRichEditControl innerControl = richEditControl.InnerControl;

            IRichEditCommandFactoryService commandFactory = innerControl.GetService<IRichEditCommandFactoryService>();
            if (commandFactory == null)
                return; // wpf richedit is not loaded

            innerControl.ReplaceService<ISyntaxHighlightService>(new SyntaxHighlightService(innerControl, codeExamplesFileExtension));

            CustomRichEditCommandFactoryService newCommandFactory = new CustomRichEditCommandFactoryService(commandFactory);
            innerControl.RemoveService(typeof(IRichEditCommandFactoryService));
            innerControl.AddService(typeof(IRichEditCommandFactoryService), newCommandFactory);

            IDocumentImportManagerService importManager = innerControl.GetService<IDocumentImportManagerService>();
            importManager.UnregisterAllImporters();
            importManager.RegisterImporter(new PlainTextDocumentImporter());
            importManager.RegisterImporter(new SourcesCodeDocumentImporter());

            IDocumentExportManagerService exportManager = innerControl.GetService<IDocumentExportManagerService>();
            exportManager.UnregisterAllExporters();
            exportManager.RegisterExporter(new PlainTextDocumentExporter());
            exportManager.RegisterExporter(new SourcesCodeDocumentExporter());

            Document document = innerControl.Document;
            document.BeginUpdate();
            try
            {
                document.DefaultCharacterProperties.FontName = "Consolas";
                document.DefaultCharacterProperties.FontSize = 10;
                document.Sections[0].Page.Width = Units.InchesToDocumentsF(20);
                document.Sections[0].Margins.Top = Units.InchesToDocumentsF(0.2f);
                document.Sections[0].Margins.Left = Units.InchesToDocumentsF(0.2f);
                document.Sections[0].Margins.Right = Units.InchesToDocumentsF(0.2f);
            }
            finally
            {
                document.EndUpdate();
            }
        }
    }
    public class SyntaxHighlightService : ISyntaxHighlightService
    {
        readonly InnerRichEditControl editor;
        readonly SyntaxHighlightInfo syntaxHighlightInfo;
        readonly string fileExtensionToHightlight;

        public SyntaxHighlightService(InnerRichEditControl editor, string extension)
        {
            this.editor = editor;

            syntaxHighlightInfo = new SyntaxHighlightInfo();
            this.fileExtensionToHightlight = extension;
        }


        void ISyntaxHighlightService.ForceExecute()
        {
            ExecuteCore();
        }
        void ISyntaxHighlightService.Execute()
        {
            ExecuteCore();
        }
        void ExecuteCore()
        {
            TokenCollection tokens = Parse(editor.Text);
            HighlightSyntax(tokens);
        }
        TokenCollection Parse(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }
            ITokenCategoryHelper tokenizer = CreateTokenizer();
            if (tokenizer == null)
            {
                return new TokenCollection();
            }
            return tokenizer.GetTokens(code);
        }

        ITokenCategoryHelper CreateTokenizer()
        {
            string fileName = editor.Options.DocumentSaveOptions.CurrentFileName;
            string extenstion;
            if (String.IsNullOrEmpty(fileName))
            {
                extenstion = this.fileExtensionToHightlight;
            }
            else
            {
                extenstion = Path.GetExtension(fileName);
            }
            ITokenCategoryHelper result = TokenCategoryHelperFactory.CreateHelperForFileExtensions(extenstion);
            if (result != null)
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        void HighlightSyntax(TokenCollection tokens)
        {
            if (tokens == null || tokens.Count == 0)
            {
                return;
            }
            Document document = editor.Document;
            CharacterProperties cp = document.BeginUpdateCharacters(0, 1);

            List<SyntaxHighlightToken> syntaxTokens = new List<SyntaxHighlightToken>(tokens.Count);
            foreach (Token token in tokens)
            {
                HighlightCategorizedToken((CategorizedToken)token, syntaxTokens);
            }
            document.ApplySyntaxHighlight(syntaxTokens);
            document.EndUpdateCharacters(cp);
        }
        void HighlightCategorizedToken(CategorizedToken token, List<SyntaxHighlightToken> syntaxTokens)
        {
            Color backColor = editor.ActiveView.BackColor;

            SyntaxHighlightProperties highlightProperties = syntaxHighlightInfo.CalculateTokenCategoryHighlight(token.Category);
            SyntaxHighlightToken syntaxToken = SetTokenColor(token, highlightProperties, backColor);
            if (syntaxToken != null)
            {
                syntaxTokens.Add(syntaxToken);
            }
        }
        SyntaxHighlightToken SetTokenColor(Token token, SyntaxHighlightProperties foreColor, Color backColor)
        {
            if (editor.Document.Paragraphs.Count < token.Range.Start.Line)
            {
                return null;
            }
            int paragraphStart = DocumentHelper.GetParagraphStart(editor.Document.Paragraphs[token.Range.Start.Line - 1]);
            int tokenStart = paragraphStart + token.Range.Start.Offset - 1;
            if (token.Range.End.Line != token.Range.Start.Line)
            {
                paragraphStart = DocumentHelper.GetParagraphStart(editor.Document.Paragraphs[token.Range.End.Line - 1]);
            }
            int tokenEnd = paragraphStart + token.Range.End.Offset - 1;
            System.Diagnostics.Debug.Assert(tokenEnd > tokenStart);
            return new SyntaxHighlightToken(tokenStart, tokenEnd - tokenStart, foreColor);
        }
    }

    public class SyntaxHighlightInfo
    {
        readonly Dictionary<TokenCategory, SyntaxHighlightProperties> properties;

        public SyntaxHighlightInfo()
        {
            properties = new Dictionary<TokenCategory, SyntaxHighlightProperties>();
            Reset();
        }
        public void Reset()
        {
            properties.Clear();
            Add(TokenCategory.Text, DXColor.Black);
            Add(TokenCategory.Keyword, DXColor.Blue);
            Add(TokenCategory.String, DXColor.Brown);
            Add(TokenCategory.Comment, DXColor.Green);
            Add(TokenCategory.Identifier, DXColor.Black);
            Add(TokenCategory.PreprocessorKeyword, DXColor.Blue);
            Add(TokenCategory.Number, DXColor.Red);
            Add(TokenCategory.Operator, DXColor.Black);
            Add(TokenCategory.Unknown, DXColor.Black);
            Add(TokenCategory.XmlComment, DXColor.Gray);

            Add(TokenCategory.CssComment, DXColor.Green);
            Add(TokenCategory.CssKeyword, DXColor.Brown);
            Add(TokenCategory.CssPropertyName, DXColor.Red);
            Add(TokenCategory.CssPropertyValue, DXColor.Blue);
            Add(TokenCategory.CssSelector, DXColor.Blue);
            Add(TokenCategory.CssStringValue, DXColor.Blue);

            Add(TokenCategory.HtmlAttributeName, DXColor.Red);
            Add(TokenCategory.HtmlAttributeValue, DXColor.Blue);
            Add(TokenCategory.HtmlComment, DXColor.Green);
            Add(TokenCategory.HtmlElementName, DXColor.Brown);
            Add(TokenCategory.HtmlEntity, DXColor.Gray);
            Add(TokenCategory.HtmlOperator, DXColor.Black);
            Add(TokenCategory.HtmlServerSideScript, DXColor.Black);
            Add(TokenCategory.HtmlString, DXColor.Blue);
            Add(TokenCategory.HtmlTagDelimiter, DXColor.Blue);
        }
        void Add(TokenCategory category, Color foreColor)
        {
            SyntaxHighlightProperties item = new SyntaxHighlightProperties();
            item.ForeColor = foreColor;
            properties.Add(category, item);
        }

        public SyntaxHighlightProperties CalculateTokenCategoryHighlight(TokenCategory category)
        {
            SyntaxHighlightProperties result = (SyntaxHighlightProperties)null;
            if (properties.TryGetValue(category, out result))
            {
                return result;
            }
            else
            {
                return properties[TokenCategory.Text];
            }
        }
    }

    public class CustomRichEditCommandFactoryService : IRichEditCommandFactoryService
    {
        readonly IRichEditCommandFactoryService service;

        public CustomRichEditCommandFactoryService(IRichEditCommandFactoryService service)
        {
            Guard.ArgumentNotNull(service, "service");
            this.service = service;
        }

        RichEditCommand IRichEditCommandFactoryService.CreateCommand(RichEditCommandId id)
        {
            if (id.Equals(RichEditCommandId.InsertColumnBreak) || id.Equals(RichEditCommandId.InsertLineBreak) || id.Equals(RichEditCommandId.InsertPageBreak))
            {
                return service.CreateCommand(RichEditCommandId.InsertParagraph);
            }
            return service.CreateCommand(id);
        }
    }

    public static class SourceCodeDocumentFormat
    {
        public static readonly DocumentFormat Id = new DocumentFormat(1325);
    }
    public class SourcesCodeDocumentImporter : PlainTextDocumentImporter
    {
        internal static readonly FileDialogFilter filter = new FileDialogFilter("Source Files", new string[] { "cs", "vb", "html", "htm", "js", "xml", "css" });
        public override FileDialogFilter Filter
        {
            get { return filter; }
        }
        public override DocumentFormat Format
        {
            get { return SourceCodeDocumentFormat.Id; }
        }
    }
    public class SourcesCodeDocumentExporter : PlainTextDocumentExporter
    {
        public override FileDialogFilter Filter
        {
            get { return SourcesCodeDocumentImporter.filter; }
        }
        public override DocumentFormat Format
        {
            get { return SourceCodeDocumentFormat.Id; }
        }
    }
}

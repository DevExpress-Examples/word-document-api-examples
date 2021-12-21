using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Import;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ImportActions
    {
        public static Action<RichEditDocumentServer> ImportRtfTextAction = ImportRtfText;
        public static Action<RichEditDocumentServer> BeforeImportAction = BeforeImport;

        static void ImportRtfText(RichEditDocumentServer wordProcessor)
        {
            #region #ImportRtfText
            // Specify the formatted text.
            string rtfString = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1049
{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}
{\f1\fswiss\fcharset0 Arial;}}
{\colortbl ;\red0\green0\blue255;}
\viewkind4\uc1\pard\cf1\lang1033\b\f0\fs32 Test.\cf0\b0\f1\fs20\par}";
            
            // Access a document.
            Document document = wordProcessor.Document;

            // Import formatted text to the document.
            document.RtfText = rtfString;
            #endregion #ImportRtfText
        }
        static void BeforeImport(RichEditDocumentServer wordProcessor)
        {
            #region #HandleBeforeImportEvent
            // Handle the Before Import event.
            wordProcessor.BeforeImport += BeforeImportHelper.BeforeImport;

            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\TerribleRevengeKOI8R.txt");            
            #endregion #HandleBeforeImportEvent
        }

        #region #@HandleBeforeImportEvent
        class BeforeImportHelper
        {
            public static void BeforeImport(object sender, BeforeImportEventArgs e)
            {
                // Specify the encoding before plain text is imported to the document.
                if (e.DocumentFormat == DocumentFormat.PlainText)
                {
                    ((PlainTextDocumentImporterOptions)e.Options).Encoding = Encoding.GetEncoding(20866);
                }
            }
        }
        #endregion #@HandleBeforeImportEvent
    }
}

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
        static void ImportRtfText(RichEditDocumentServer server)
        {
            #region #ImportRtfText
            string rtfString = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1049
{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}
{\f1\fswiss\fcharset0 Arial;}}
{\colortbl ;\red0\green0\blue255;}
\viewkind4\uc1\pard\cf1\lang1033\b\f0\fs32 Test.\cf0\b0\f1\fs20\par}";
            Document document = server.Document;
            document.RtfText = rtfString;
            #endregion #ImportRtfText
        }
        static void BeforeImport(RichEditDocumentServer server)
        {
            #region #HandleBeforeImportEvent
            server.BeforeImport += BeforeImportHelper.BeforeImport;
            server.LoadDocument("Documents\\TerribleRevengeKOI8R.txt");            
            #endregion #HandleBeforeImportEvent
        }

        #region #@HandleBeforeImportEvent
        class BeforeImportHelper
        {
            public static void BeforeImport(object sender, BeforeImportEventArgs e)
            {
                if (e.DocumentFormat == DocumentFormat.PlainText)
                {
                    ((PlainTextDocumentImporterOptions)e.Options).Encoding = Encoding.GetEncoding(20866);
                }
            }
        }
        #endregion #@HandleBeforeImportEvent
    }
}

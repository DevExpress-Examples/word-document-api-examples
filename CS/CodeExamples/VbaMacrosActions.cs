using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class VbaMacrosActions
    {
        public static Action<RichEditDocumentServer> ObtainVbaMacrosAction = ObtainVbaMacros;
        public static Action<RichEditDocumentServer> ClearVbaModulesAction = ClearVbaModules;

        static void ObtainVbaMacros(RichEditDocumentServer wordProcessor)
        {
            #region #ObtainVbaMacros
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx");
            if (document.VbaProject.Modules.Count > 0)
                foreach (VbaModule module in document.VbaProject.Modules)
                { document.AppendText("\r\n \u00B7 " + module.Name); }

            #endregion #ObtainVbaMacros
        }

        static void ClearVbaModules(RichEditDocumentServer wordProcessor)
        {
            #region #ClearVbaModules
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx");
            if (document.VbaProject.Modules.Count > 0)
                document.VbaProject.Modules.Clear();
            #endregion #ClearVbaModules
        }

    }
}

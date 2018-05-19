using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class DocumentPropertiesActions
    {
        static void StandardDocumentProperties(RichEditDocumentServer server)
        {
            #region #StandardDocumentProperties
            server.CreateNewDocument();
            Document document = server.Document;
            document.BeginUpdate();

            document.DocumentProperties.Creator = "John Doe";
            document.DocumentProperties.Title = "Inserting Custom Properties";
            document.DocumentProperties.Category = "TestDoc";
            document.DocumentProperties.Description = "This code demonstrates API to modify and display standard document properties.";

            document.Fields.Create(document.AppendText("\nAUTHOR: ").End, "AUTHOR");
            document.Fields.Create(document.AppendText("\nTITLE: ").End, "TITLE");
            document.Fields.Create(document.AppendText("\nCOMMENTS: ").End, "COMMENTS");
            document.Fields.Create(document.AppendText("\nCREATEDATE: ").End, "CREATEDATE");
            document.Fields.Create(document.AppendText("\nCategory: ").End, "DOCPROPERTY Category");
            document.Fields.Update();
            document.EndUpdate();
            #endregion #StandardDocumentProperties
        }
        static void CustomDocumentProperties(RichEditDocumentServer server)
        {
            #region #CustomDocumentProperties
            server.CreateNewDocument();
            Document document = server.Document;
            document.BeginUpdate();
            document.Fields.Create(document.AppendText("\nMyNumericProperty: ").End, "DOCVARIABLE CustomProperty MyNumericProperty");
            document.Fields.Create(document.AppendText("\nMyStringProperty: ").End, "DOCVARIABLE CustomProperty MyStringProperty");
            document.Fields.Create(document.AppendText("\nMyBooleanProperty: ").End, "DOCVARIABLE CustomProperty MyBooleanProperty");
            document.EndUpdate();

            document.CustomProperties.Add("MyNumericProperty", 123.45);
            document.CustomProperties.Add("MyStringProperty", "The Final Answer");
            document.CustomProperties.Add("MyBooleanProperty", true);

            server.CalculateDocumentVariable += DocumentPropertyDisplayHelper.OnCalculateDocumentVariable;
            document.Fields.Update();
            #endregion #CustomDocumentProperties
        }

        #region #@CustomDocumentProperties
        class DocumentPropertyDisplayHelper
        {
           public static void OnCalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
            {
                if (e.Arguments.Count == 0 || e.VariableName != "CustomProperty")
                    return;

                string name = e.Arguments[0].Value;
                object customProperty = ((RichEditDocumentServer)sender).Document.CustomProperties[name];
                if (customProperty != null)
                    e.Value = customProperty.ToString();
                e.Handled = true;
            }
        }
        #endregion #@CustomDocumentProperties
    }
}

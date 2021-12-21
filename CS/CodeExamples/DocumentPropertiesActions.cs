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
        public static Action<RichEditDocumentServer> StandardDocumentPropertiesAction = StandardDocumentProperties;
        public static Action<RichEditDocumentServer> CustomDocumentPropertiesAction = CustomDocumentProperties;

        static void StandardDocumentProperties(RichEditDocumentServer wordProcessor)
        {
            #region #StandardDocumentProperties
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Set the built-in document properties.
            document.DocumentProperties.Creator = "John Doe";
            document.DocumentProperties.Title = "Inserting Custom Properties";
            document.DocumentProperties.Category = "TestDoc";
            document.DocumentProperties.Description = "This code demonstrates API to modify and display standard document properties.";

            // Display the specified built-in properties in the document.
            document.Fields.Create(document.AppendText("\nAUTHOR: ").End, "AUTHOR");
            document.Fields.Create(document.AppendText("\nTITLE: ").End, "TITLE");
            document.Fields.Create(document.AppendText("\nCOMMENTS: ").End, "COMMENTS");
            document.Fields.Create(document.AppendText("\nCREATEDATE: ").End, "CREATEDATE");
            document.Fields.Create(document.AppendText("\nCategory: ").End, "DOCPROPERTY Category");
            document.Fields.Update();
            
            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #StandardDocumentProperties
        }


        static void CustomDocumentProperties(RichEditDocumentServer wordProcessor)
        {
            #region #CustomDocumentProperties
            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Display the custom document properties in the document.
            document.Fields.Create(document.AppendText("\nMyNumericProperty: ").End, "DOCVARIABLE CustomProperty MyNumericProperty");
            document.Fields.Create(document.AppendText("\nMyStringProperty: ").End, "DOCVARIABLE CustomProperty MyStringProperty");
            document.Fields.Create(document.AppendText("\nMyBooleanProperty: ").End, "DOCVARIABLE CustomProperty MyBooleanProperty");
            
            // Finalize to edit the document.
            document.EndUpdate();

            // Set the custom document properties.
            document.CustomProperties["MyNumericProperty"] = 123.45;
            document.CustomProperties["MyStringProperty"] = "The Final Answer";
            document.CustomProperties["MyBooleanProperty"] = true;

            wordProcessor.CalculateDocumentVariable += DocumentPropertyDisplayHelper.OnCalculateDocumentVariable;
            
            // Update all fields in the main document body.
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

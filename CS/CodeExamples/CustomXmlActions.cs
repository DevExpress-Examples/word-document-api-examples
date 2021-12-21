using DevExpress.XtraRichEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class CustomXmlActions
    {
        public static Action<RichEditDocumentServer> AddCustomXmlPartAction = AddCustomXmlPart;
        public static Action<RichEditDocumentServer> AccessCustomXmlPartAction = AccessCustomXmlPart;
        public static Action<RichEditDocumentServer> RemoveCustomXmlPartAction = RemoveCustomXmlPart;
        static void AddCustomXmlPart(RichEditDocumentServer wordProcessor)
        {
            #region #AddCustomXmlPart
            // Access a document.
            Document document = wordProcessor.Document;
            
            // Append text to the document.
            document.AppendText("This document contains custom XML parts.");
            
            // Add an empty custom XML part.
            ICustomXmlPart xmlItem = document.CustomXmlParts.Add();
            
            // Populate the XML part with content.
            XmlElement elem = xmlItem.CustomXmlPartDocument.CreateElement("Employees");
            elem.InnerText = "Stephen Edwards";
            xmlItem.CustomXmlPartDocument.AppendChild(elem);

            // Specify the custom XML part content.
            string xmlString = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Stephen</FirstName>
                                <LastName>Edwards</LastName>
                                <Address>4726 - 11th Ave. N.E.</Address>
                                <City>Seattle</City>
                                <Region>WA</Region>
                                <PostalCode>98122</PostalCode>
                                <Country>USA</Country>
                            </Employees>";
            document.CustomXmlParts.Insert(1, xmlString);

            // Add a custom XML part from a file.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("Documents\\Employees.xml");
            document.CustomXmlParts.Add(xmlDoc);
            #endregion #AddCustomXmlPart
        }

        static void AccessCustomXmlPart(RichEditDocumentServer wordProcessor)
        {
            #region #AccessCustomXmlPart
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.CustomXmlParts.Count > 0)
            {
                // Access a custom XML file stored in the document.
                XmlDocument xmlDoc = document.CustomXmlParts[0].CustomXmlPartDocument;

                // Retrieve employee names from the XML file and display them in the document.
                XmlNodeList nameList = xmlDoc.GetElementsByTagName("Name");
                document.AppendText("Employee list:");
                foreach (XmlNode name in nameList)
                {
                    document.AppendText("\r\n \u00B7 " + name.InnerText);
                }
            }
            #endregion #AccessCustomXmlPart
        }

        static void RemoveCustomXmlPart(RichEditDocumentServer wordProcessor)
        {
            #region #RemoveCustomXmlPart
            // Access a document.
            Document document = wordProcessor.Document;

            // Append text to the document.
            document.AppendText("This document contains custom XML parts.");

            // Add the first custom XML part.
            string xmlString1 = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Stephen</FirstName>
                                <LastName>Edwards</LastName>
                            </Employees>";
            var xmlItem1 = document.CustomXmlParts.Add(xmlString1);

            // Add the second custom XML part.
            string xmlString2 = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                            <Employees>
                                <FirstName>Andrew</FirstName>
                                <LastName>Fuller</LastName>
                            </Employees>";
            var xmlItem2 = document.CustomXmlParts.Add(xmlString2);

            // Remove the first item from the collection.
            document.CustomXmlParts.Remove(xmlItem1);
            // Use the RemoveAt method to remove an item at the specified position from the collection.
            // document.CustomXmlParts.RemoveAt(0);
            // Use the Clear method to remove all items from the collection.
            // document.CustomXmlParts.Clear();
            #endregion #RemoveCustomXmlPart
        }
    }
}

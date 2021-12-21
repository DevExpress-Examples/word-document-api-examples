Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module DocumentPropertiesActions

        Public StandardDocumentPropertiesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.DocumentPropertiesActions.StandardDocumentProperties

        Public CustomDocumentPropertiesAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.DocumentPropertiesActions.CustomDocumentProperties

        Private Sub StandardDocumentProperties(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#StandardDocumentProperties"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Set the built-in document properties.
            document.DocumentProperties.Creator = "John Doe"
            document.DocumentProperties.Title = "Inserting Custom Properties"
            document.DocumentProperties.Category = "TestDoc"
            document.DocumentProperties.Description = "This code demonstrates API to modify and display standard document properties."
            ' Display the specified built-in properties in the document.
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "AUTHOR: "))).[End], "AUTHOR")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "TITLE: "))).[End], "TITLE")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "COMMENTS: "))).[End], "COMMENTS")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "CREATEDATE: "))).[End], "CREATEDATE")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "Category: "))).[End], "DOCPROPERTY Category")
            document.Fields.Update()
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #StandardDocumentProperties
        End Sub

        Private Sub CustomDocumentProperties(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CustomDocumentProperties"
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Display the custom document properties in the document.
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "MyNumericProperty: "))).[End], "DOCVARIABLE CustomProperty MyNumericProperty")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "MyStringProperty: "))).[End], "DOCVARIABLE CustomProperty MyStringProperty")
            document.Fields.Create(document.AppendText(CStr((Global.Microsoft.VisualBasic.Constants.vbLf & "MyBooleanProperty: "))).[End], "DOCVARIABLE CustomProperty MyBooleanProperty")
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Set the custom document properties.
            document.CustomProperties("MyNumericProperty") = 123.45
            document.CustomProperties("MyStringProperty") = "The Final Answer"
            document.CustomProperties("MyBooleanProperty") = True
            AddHandler wordProcessor.CalculateDocumentVariable, AddressOf RichEditDocumentServerAPIExample.CodeExamples.DocumentPropertiesActions.DocumentPropertyDisplayHelper.OnCalculateDocumentVariable
            ' Update all fields in the main document body.
            document.Fields.Update()
#End Region  ' #CustomDocumentProperties
        End Sub

#Region "#@CustomDocumentProperties"
        Private Class DocumentPropertyDisplayHelper

            Public Shared Sub OnCalculateDocumentVariable(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs)
                If e.Arguments.Count = 0 OrElse Not Equals(e.VariableName, "CustomProperty") Then Return
                Dim name As String = e.Arguments(CInt((0))).Value
                Dim customProperty As Object = CType(sender, DevExpress.XtraRichEdit.RichEditDocumentServer).Document.CustomProperties(name)
                If customProperty IsNot Nothing Then e.Value = customProperty.ToString()
                e.Handled = True
            End Sub
        End Class
#End Region  ' #@CustomDocumentProperties
    End Module
End Namespace

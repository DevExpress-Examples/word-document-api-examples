Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Public Module DocumentPropertiesActions
		Private Sub StandardDocumentProperties(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#StandardDocumentProperties"
			wordProcessor.CreateNewDocument()
			Dim document As Document = wordProcessor.Document
			document.BeginUpdate()

			document.DocumentProperties.Creator = "John Doe"
			document.DocumentProperties.Title = "Inserting Custom Properties"
			document.DocumentProperties.Category = "TestDoc"
			document.DocumentProperties.Description = "This code demonstrates API to modify and display standard document properties."

			document.Fields.Create(document.AppendText(vbLf & "AUTHOR: ").End, "AUTHOR")
			document.Fields.Create(document.AppendText(vbLf & "TITLE: ").End, "TITLE")
			document.Fields.Create(document.AppendText(vbLf & "COMMENTS: ").End, "COMMENTS")
			document.Fields.Create(document.AppendText(vbLf & "CREATEDATE: ").End, "CREATEDATE")
			document.Fields.Create(document.AppendText(vbLf & "Category: ").End, "DOCPROPERTY Category")
			document.Fields.Update()
			document.EndUpdate()
'			#End Region ' #StandardDocumentProperties
		End Sub


		Private Sub CustomDocumentProperties(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#CustomDocumentProperties"
			wordProcessor.CreateNewDocument()
			Dim document As Document = wordProcessor.Document
			document.BeginUpdate()
			document.Fields.Create(document.AppendText(vbLf & "MyNumericProperty: ").End, "DOCVARIABLE CustomProperty MyNumericProperty")
			document.Fields.Create(document.AppendText(vbLf & "MyStringProperty: ").End, "DOCVARIABLE CustomProperty MyStringProperty")
			document.Fields.Create(document.AppendText(vbLf & "MyBooleanProperty: ").End, "DOCVARIABLE CustomProperty MyBooleanProperty")
			document.EndUpdate()

			document.CustomProperties("MyNumericProperty") = 123.45
			document.CustomProperties("MyStringProperty") = "The Final Answer"
			document.CustomProperties("MyBooleanProperty") = True

			AddHandler wordProcessor.CalculateDocumentVariable, AddressOf DocumentPropertyDisplayHelper.OnCalculateDocumentVariable
			document.Fields.Update()
'			#End Region ' #CustomDocumentProperties
		End Sub

		#Region "#@CustomDocumentProperties"
		Private Class DocumentPropertyDisplayHelper
		   Public Shared Sub OnCalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
				If e.Arguments.Count = 0 OrElse e.VariableName <> "CustomProperty" Then
					Return
				End If

				Dim name As String = e.Arguments(0).Value
				Dim customProperty As Object = DirectCast(sender, RichEditDocumentServer).Document.CustomProperties(name)
				If customProperty IsNot Nothing Then
					e.Value = customProperty.ToString()
				End If
				e.Handled = True
		   End Sub
		End Class
		#End Region ' #@CustomDocumentProperties


	End Module
End Namespace

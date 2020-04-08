Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Import

Namespace RichEditDocumentServerAPIExample.CodeExamples
	Friend Class ImportActions
		Private Shared Sub ImportRtfText(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#ImportRtfText"
			Dim rtfString As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1049
{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}
{\f1\fswiss\fcharset0 Arial;}}
{\colortbl ;\red0\green0\blue255;}
\viewkind4\uc1\pard\cf1\lang1033\b\f0\fs32 Test.\cf0\b0\f1\fs20\par}"
			Dim document As Document = wordProcessor.Document
			document.RtfText = rtfString
'			#End Region ' #ImportRtfText
		End Sub
		Private Shared Sub BeforeImport(ByVal wordProcessor As RichEditDocumentServer)
'			#Region "#HandleBeforeImportEvent"
			AddHandler wordProcessor.BeforeImport, AddressOf BeforeImportHelper.BeforeImport
			wordProcessor.LoadDocument("Documents\TerribleRevengeKOI8R.txt")
'			#End Region ' #HandleBeforeImportEvent
		End Sub

		#Region "#@HandleBeforeImportEvent"
		Private Class BeforeImportHelper
			Public Shared Sub BeforeImport(ByVal sender As Object, ByVal e As BeforeImportEventArgs)
				If e.DocumentFormat = DocumentFormat.PlainText Then
					CType(e.Options, PlainTextDocumentImporterOptions).Encoding = Encoding.GetEncoding(20866)
				End If
			End Sub
		End Class
		#End Region ' #@HandleBeforeImportEvent
	End Class
End Namespace

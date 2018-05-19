Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Import

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Public NotInheritable Class ImportActions

        Private Sub New()
        End Sub

        Private Shared Sub ImportRtfText(ByVal server As RichEditDocumentServer)
'            #Region "#ImportRtfText"
            Dim rtfString As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1049" & ControlChars.CrLf & _
"{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}" & ControlChars.CrLf & _
"{\f1\fswiss\fcharset0 Arial;}}" & ControlChars.CrLf & _
"{\colortbl ;\red0\green0\blue255;}" & ControlChars.CrLf & _
"\viewkind4\uc1\pard\cf1\lang1033\b\f0\fs32 Test.\cf0\b0\f1\fs20\par}"
            Dim document As Document = server.Document
            document.RtfText = rtfString
'            #End Region ' #ImportRtfText
        End Sub
        Private Shared Sub BeforeImport(ByVal server As RichEditDocumentServer)
'            #Region "#HandleBeforeImportEvent"
            server.LoadDocument("Documents\TerribleRevengeKOI8R.txt")
            AddHandler server.BeforeImport, AddressOf BeforeImportHelper.BeforeImport
'            #End Region ' #HandleBeforeImportEvent
        End Sub

        #Region "#@HandleBeforeImportEvent"
        Private Class BeforeImportHelper
            Public Shared Sub BeforeImport(ByVal sender As Object, ByVal e As BeforeImportEventArgs)
                If e.DocumentFormat = DocumentFormat.PlainText Then
                    CType(e.Options, PlainTextDocumentImporterOptions).Encoding = Encoding.UTF32
                End If
            End Sub
        End Class
        #End Region ' #@HandleBeforeImportEvent
    End Class
End Namespace

Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module VbaMacrosActions

        Public ObtainVbaMacrosAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ObtainVbaMacros

        Public ClearVbaModulesAction As Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf ClearVbaModules

        Private Sub ObtainVbaMacros(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ObtainVbaMacros"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx")
            If document.VbaProject.Modules.Count > 0 Then
                For Each [module] As DevExpress.XtraRichEdit.API.Native.VbaModule In document.VbaProject.Modules
                    document.AppendText(Global.Microsoft.VisualBasic.Constants.vbCrLf & " · " & [module].Name)
                Next
            End If
#End Region  ' #ObtainVbaMacros
        End Sub

        Private Sub ClearVbaModules(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ClearVbaModules"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents\Grimm.docx")
            If document.VbaProject.Modules.Count > 0 Then document.VbaProject.Modules.Clear()
#End Region  ' #ClearVbaModules
        End Sub
    End Module
End Namespace

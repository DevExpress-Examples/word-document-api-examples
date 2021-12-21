Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module ProtectionActions

        Private Sub ProtectDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#ProtectDocument"
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If Not document.IsDocumentProtected Then
                'Protect the document with a password
                document.Protect("123", DevExpress.XtraRichEdit.API.Native.DocumentProtectionType.[ReadOnly])
                'Insert a comment indicating that the document is protected
                document.Comments.Create(document.Paragraphs(CInt((0))).Range, "Admin")
                Dim commentDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Comments(CInt((0))).BeginUpdate()
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is protected with a password." & Global.Microsoft.VisualBasic.Constants.vbLf & "You cannot modify the document until protection is removed.")
                commentDocument.EndUpdate()
            End If
#End Region  ' #ProtectDocument
        End Sub

        Private Sub UnprotectDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#UnprotectDocument"
            wordProcessor.LoadDocument("Documents//Grimm_Protected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            If document.IsDocumentProtected = True Then
                'Unprotect the document
                document.Unprotect()
                'Insert a comment indicating that the document can be edited
                document.Comments.Create(document.Paragraphs(CInt((0))).Range, "Admin")
                Dim commentDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Comments(CInt((0))).BeginUpdate()
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is unprotected. You can modify the document according to your requests.")
                commentDocument.EndUpdate()
            End If
#End Region  ' #UnprotectDocument
        End Sub

        Private Sub CreateRangePermissions(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateRangePermissions"
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Protect document range
            Dim rangePermissions As DevExpress.XtraRichEdit.API.Native.RangePermissionCollection = document.BeginUpdateRangePermissions()
            Dim rp As DevExpress.XtraRichEdit.API.Native.RangePermission = rangePermissions.CreateRangePermission(document.Paragraphs(CInt((3))).Range)
            rp.Group = "Administrators"
            rp.UserName = "admin@somecompany.com"
            rangePermissions.Add(rp)
            document.EndUpdateRangePermissions(rangePermissions)
            ' Enforce protection and set password.
            document.Protect("123")
#End Region  ' #CreateRangePermissions
        End Sub
    End Module
End Namespace

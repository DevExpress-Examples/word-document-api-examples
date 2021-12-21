Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Public Module ProtectionActions

        Public ProtectDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ProtectionActions.ProtectDocument

        Public UnprotectDocumentAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ProtectionActions.UnprotectDocument

        Public CreateRangePermissionsAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ProtectionActions.CreateRangePermissions

        Private Sub ProtectDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
            Call RichEditDocumentServerAPIExample.CodeExamples.ProtectionActions.UnprotectResultingDocument(wordProcessor)
#Region "#ProtectDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check whether the document is protected.
            If Not document.IsDocumentProtected Then
                ' Protect the document with a password.
                document.Protect("123", DevExpress.XtraRichEdit.API.Native.DocumentProtectionType.[ReadOnly])
                ' Create a comment related to the first paragraph.
                document.Comments.Create(document.Paragraphs(CInt((0))).Range, "Admin")
                ' Access the comment content.
                Dim commentDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Comments(CInt((0))).BeginUpdate()
                ' Specify the comment text to indicate that the document is protected.
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is protected with a password." & Global.Microsoft.VisualBasic.Constants.vbLf & "You cannot modify the document until protection is removed.")
                ' Finalize to edit the comment.
                commentDocument.EndUpdate()
                ' Save and open the protected document.
                wordProcessor.SaveDocument("ResultProtected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start("ResultProtected.docx")
            End If
#End Region  ' #ProtectDocument
        End Sub

        Private Sub UnprotectDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#UnprotectDocument"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm_Protected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Check whether the document is protected.
            If document.IsDocumentProtected = True Then
                ' Unprotect the document.
                document.Unprotect()
                ' Create a comment related to the first paragraph.
                document.Comments.Create(document.Paragraphs(CInt((0))).Range, "Admin")
                ' Access the comment content.
                Dim commentDocument As DevExpress.XtraRichEdit.API.Native.SubDocument = document.Comments(CInt((0))).BeginUpdate()
                ' Specify the comment text to indicate that the document is unprotected.
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is unprotected. You can modify the document according to your requests.")
                ' Finalize to edit the comment.
                commentDocument.EndUpdate()
                ' Save and open the protected document.
                wordProcessor.SaveDocument("ResultUnrotected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
                System.Diagnostics.Process.Start("ResultUnprotected.docx")
            End If
#End Region  ' #UnprotectDocument
        End Sub

        Private Sub CreateRangePermissions(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
            Call RichEditDocumentServerAPIExample.CodeExamples.ProtectionActions.UnprotectResultingDocument(wordProcessor)
#Region "#CreateRangePermissions"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Access the range permissions collection.
            Dim rangePermissions As DevExpress.XtraRichEdit.API.Native.RangePermissionCollection = document.BeginUpdateRangePermissions()
            If document.Paragraphs.Count > 3 Then
                ' Specify the group of users and the user that are allowed to edit the document range.
                Dim rp As DevExpress.XtraRichEdit.API.Native.RangePermission = rangePermissions.CreateRangePermission(document.Paragraphs(CInt((3))).Range)
                rp.Group = "Administrators"
                rp.UserName = "admin@somecompany.com"
                rangePermissions.Add(rp)
            End If

            ' Finalize to update the range permissions collection.
            document.EndUpdateRangePermissions(rangePermissions)
            ' Protect the document with a password.
            document.Protect("123")
            ' Save and open the protected document.
            wordProcessor.SaveDocument("ResultProtected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            System.Diagnostics.Process.Start("ResultProtected.docx")
#End Region  ' #CreateRangePermissions
        End Sub

        Private Sub UnprotectResultingDocument(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
            Try
                ' Load a document from a file.
                wordProcessor.LoadDocument("ResultProtected.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
                ' Access a document.
                Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
                If document.IsDocumentProtected = True Then
                    ' Unprotect the document.
                    document.Unprotect()
                End If
            Catch
            End Try
        End Sub
    End Module
End Namespace

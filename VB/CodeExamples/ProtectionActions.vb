﻿Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Public NotInheritable Class ProtectionActions

        Private Sub New()
        End Sub

        Private Shared Sub ProtectDocument(ByVal server As RichEditDocumentServer)
'            #Region "#ProtectDocument"
            server.LoadDocument("Documents//Grimm.docx",DocumentFormat.OpenXml)
            Dim document As Document = server.Document
            If Not document.IsDocumentProtected Then
                'Protect the document with a password
                document.Protect("123",DocumentProtectionType.ReadOnly)

                'Insert a comment indicating that the document is protected
                document.Comments.Create(document.Paragraphs(0).Range, "Admin")
                Dim commentDocument As SubDocument = document.Comments(0).BeginUpdate()
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is protected with a password." & vbLf & "You cannot modify the document until protection is removed.")
                commentDocument.EndUpdate()
            End If
'            #End Region ' #ProtectDocument
        End Sub
        Private Shared Sub UnprotectDocument(ByVal server As RichEditDocumentServer)
'            #Region "#UnprotectDocument"
            server.LoadDocument("Documents//Grimm_Protected.docx", DocumentFormat.OpenXml)
            Dim document As Document = server.Document

            If document.IsDocumentProtected = True Then
                'Unprotect the document
                document.Unprotect()

                'Insert a comment indicating that the document can be edited
                document.Comments.Create(document.Paragraphs(0).Range,"Admin")
                Dim commentDocument As SubDocument = document.Comments(0).BeginUpdate()
                commentDocument.InsertText(commentDocument.CreatePosition(0), "Document is unprotected. You can modify the document according to your requests.")
                commentDocument.EndUpdate()
            End If
'            #End Region ' #UnprotectDocument
        End Sub
        Private Shared Sub CreateRangePermissions(ByVal server As RichEditDocumentServer)
'            #Region "#CreateRangePermissions"
            server.LoadDocument("Documents//Grimm.docx", DocumentFormat.OpenXml)
            Dim document As Document = server.Document

            ' Protect document range
            Dim rangePermissions As RangePermissionCollection = document.BeginUpdateRangePermissions()
            Dim rp As RangePermission = rangePermissions.CreateRangePermission(document.Paragraphs(3).Range)
            rp.Group = "Everyone"
            rangePermissions.Add(rp)

            document.EndUpdateRangePermissions(rangePermissions)
            ' Enforce protection and set password.
            document.Protect("123")
'            #End Region ' #CreateRangePermissions
        End Sub
    End Class
End Namespace

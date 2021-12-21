Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class ListsActions

        Public Shared CreateBulletedListAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ListsActions.CreateBulletedList

        Public Shared CreateNumberedListAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ListsActions.CreateNumberedList

        Public Shared CreateMultilevelListAction As System.Action(Of DevExpress.XtraRichEdit.RichEditDocumentServer) = AddressOf RichEditDocumentServerAPIExample.CodeExamples.ListsActions.CreateMultilevelList

        Private Shared Sub CreateBulletedList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateBulletedList"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a new list pattern object.
            Dim list As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            ' Specify the list type.
            list.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.Bullet
            ' Access the first list level.
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = list.Levels(0)
            ' Specify the left indent of the level's paragraph.
            level.ParagraphProperties.LeftIndent = 100
            ' Specify the format of bullets.
            ' Without this step, the list is considered as numbered.
            level.DisplayFormatString = "·"
            level.CharacterProperties.FontName = "Symbol"
            ' Create a new list based on the specified pattern.
            Dim bulletedList As DevExpress.XtraRichEdit.API.Native.NumberingList = document.NumberingLists.Add(0)
            ' Access the collection of paragraphs.
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            ' Apply the numbering list format to the document paragraphs.
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0)
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #CreateBulletedList
        End Sub

        Private Shared Sub CreateNumberedList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNumberedList"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a new list pattern object.
            Dim abstractListNumberingRoman As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            ' Specify the list type.
            abstractListNumberingRoman.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.Simple
            ' Specify properties of the first list level.
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = abstractListNumberingRoman.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            ' Specify the numbering style for the list level.
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerRoman
            level.DisplayFormatString = "{0}."
            ' Create a new list based on the specified pattern.
            Dim numberingList As DevExpress.XtraRichEdit.API.Native.NumberingList = document.NumberingLists.Add(0)
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Start to edit the document.
            document.BeginUpdate()
            ' Access the collection of paragraphs.
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            ' Apply the numbering list format to the document paragraphs.
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0)
            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #CreateNumberedList
        End Sub

        Private Shared Sub CreateMultilevelList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateMultilevelList"
            ' Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            ' Access a document.
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            ' Start to edit the document.
            document.BeginUpdate()
            ' Create a new list pattern object.
            Dim list As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            ' Specify the list type.
            list.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.MultiLevel
            ' Specify parameters for the first list level.
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 105
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 55
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.UpperRoman
            level.DisplayFormatString = "{0}"
            ' Specify parameters for the second list level.
            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 125
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 65
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerRoman
            level.DisplayFormatString = "{1})"
            ' Specify parameters for the third list level.
            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 145
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerLetter
            level.DisplayFormatString = "{2}."
            ' Create a new list based on the specified pattern.
            document.NumberingLists.Add(0)
            ' Finalize to edit the document.
            document.EndUpdate()
            ' Start to edit the document.
            document.BeginUpdate()
            ' Convert all paragraphs to list items.
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            For Each pgf As DevExpress.XtraRichEdit.API.Native.Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = pgf.Index
            Next

            ' Finalize to edit the document.
            document.EndUpdate()
#End Region  ' #CreateMultilevelList
        End Sub
    End Class
End Namespace

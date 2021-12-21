Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples

    Friend Class ListsActions

        Private Shared Sub CreateBulletedList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateBulletedList"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents//List.docx")
            document.BeginUpdate()
            ' Create a new list pattern object
            Dim list As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            'Specify the list's type
            list.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.Bullet
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 100
            'Specify the bullets' format
            'Without this step, the list is considered as numbered
            level.DisplayFormatString = "·"
            level.CharacterProperties.FontName = "Symbol"
            'Create a new list based on the specific pattern
            Dim bulletedList As DevExpress.XtraRichEdit.API.Native.NumberingList = document.NumberingLists.Add(0)
            ' Add paragraphs to the list
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0)
            document.EndUpdate()
#End Region  ' #CreateBulletedList
        End Sub

        Private Shared Sub CreateNumberedList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateNumberedList"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents//List.docx")
            document.BeginUpdate()
            'Create a new pattern object
            Dim abstractListNumberingRoman As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            'Specify the list's type
            abstractListNumberingRoman.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.Simple
            'Define the first level's properties
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = abstractListNumberingRoman.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            'Specify the roman format
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerRoman
            level.DisplayFormatString = "{0}."
            'Create a new list based on the specific pattern
            Dim numberingList As DevExpress.XtraRichEdit.API.Native.NumberingList = document.NumberingLists.Add(0)
            document.EndUpdate()
            document.BeginUpdate()
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            'Add paragraphs to the list
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0)
            document.EndUpdate()
#End Region  ' #CreateNumberedList
        End Sub

        Private Shared Sub CreateMultilevelList(ByVal wordProcessor As DevExpress.XtraRichEdit.RichEditDocumentServer)
#Region "#CreateMultilevelList"
            Dim document As DevExpress.XtraRichEdit.API.Native.Document = wordProcessor.Document
            document.LoadDocument("Documents//List.docx")
            document.BeginUpdate()
            'Create a new pattern object
            Dim list As DevExpress.XtraRichEdit.API.Native.AbstractNumberingList = document.AbstractNumberingLists.Add()
            'Specify the list's type
            list.NumberingType = DevExpress.XtraRichEdit.API.Native.NumberingType.MultiLevel
            'Specify parameters for each list level
            Dim level As DevExpress.XtraRichEdit.API.Native.ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 105
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 55
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.UpperRoman
            level.DisplayFormatString = "{0}"
            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 125
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 65
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerRoman
            level.DisplayFormatString = "{1})"
            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 145
            level.ParagraphProperties.FirstLineIndentType = DevExpress.XtraRichEdit.API.Native.ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = DevExpress.XtraRichEdit.API.Native.NumberingFormat.LowerLetter
            level.DisplayFormatString = "{2}."
            'Create a new list object based on the specified pattern
            document.NumberingLists.Add(0)
            document.EndUpdate()
            'Convert all paragraphs to list items
            document.BeginUpdate()
            Dim paragraphs As DevExpress.XtraRichEdit.API.Native.ParagraphCollection = document.Paragraphs
            For Each pgf As DevExpress.XtraRichEdit.API.Native.Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = pgf.Index
            Next

            document.EndUpdate()
#End Region  ' #CreateMultilevelList
        End Sub
    End Class
End Namespace

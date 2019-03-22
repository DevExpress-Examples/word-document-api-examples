Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditDocumentServerAPIExample.CodeExamples
    Friend Class ListsActions

        Private Shared Sub CreateBulletedList(ByVal server As RichEditDocumentServer)
            '            #Region "#CreateBulletedList
            Dim document As Document = server.Document
            document.LoadDocument("Documents//List.docx")
            document.BeginUpdate()

            ' Create a new list pattern objects
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()

            ' Specify the list's type
            list.NumberingType = NumberingType.Bullet
            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 100

            ' Specify the bullet's format
            ' Without this step, the list is considered as numbered
            level.DisplayFormatString = "·"
            level.CharacterProperties.FontName = "Symbol"

            ' Create a new list based on the specific pattern
            Dim bulletedList As NumberingList = document.NumberingLists.Add(0)

            ' Add paragraphs to the list
            Dim paragraphs As ParagraphCollection = document.Paragraphs
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0)
            document.EndUpdate()
            '            #End Region ' #CreateBulletedList
        End Sub


        Private Shared Sub CreateNumberedList(ByVal server As RichEditDocumentServer)
            '            #Region "#CreateNumberedList
            Dim document As Document = server.Document
            document.LoadDocument("Documents//List.docx")
            document.BeginUpdate()

            'Create a new pattern object
            Dim abstractListNumberingRoman As AbstractNumberingList = document.AbstractNumberingLists.Add()

            'Specify the list's type
            abstractListNumberingRoman.NumberingType = NumberingType.Simple

            'Define the first level's properties
            Dim level As ListLevel = abstractListNumberingRoman.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerRoman
            level.DisplayFormatString = "{0}."

            'Create a new list based on the specific pattern
            Dim numberingList As NumberingList = document.NumberingLists.Add(0)
            document.EndUpdate()

            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs

            'Add paragraphs to the list
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0)
            document.EndUpdate()
            '            #End Region ' #CreateNumberedList
        End Sub

        Private Shared Sub CreateMultilevelList(ByVal server As RichEditDocumentServer)
            ' #Region "#CreateMultilevelList 
            Dim document As Document = server.Document
            document.LoadDocument("Documents//List.docx")

            document.BeginUpdate()

            'Create a new list pattern object
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()

            'Specify the list's type
            list.NumberingType = NumberingType.MultiLevel

            'Specify parameters for each level
            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 105
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 55
            level.Start = 1
            level.NumberingFormat = NumberingFormat.UpperRoman
            level.DisplayFormatString = "{0}"

            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 125
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 65
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerRoman
            level.DisplayFormatString = "{1})"

            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 145
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerLetter
            level.DisplayFormatString = "{2}."

            'Create a new list object based on the specified pattern
            document.NumberingLists.Add(0)
            document.EndUpdate()

            document.BeginUpdate()

            ' Apply numbering to a list
            Dim paragraphs As ParagraphCollection = document.Paragraphs

            For Each pgf As Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = pgf.Index
            Next

            document.EndUpdate()
            ' #End Region '#CreateMultilevelList
        End Sub

    End Class
End Namespace


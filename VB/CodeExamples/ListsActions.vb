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
            document.BeginUpdate()
            ' Define an abstract list that is the pattern for lists used in the document.
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()
            list.NumberingType = NumberingType.Bullet

            ' Specify parameters for each list level.

            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 100
            level.CharacterProperties.FontName = "Symbol"
            level.DisplayFormatString = New String(ChrW(&HB7), 1)


            ' Create a list for use in the document. It is based on a previously defined abstract list with ID = 0.
            Dim bulletedList As NumberingList = document.NumberingLists.Add(0)
            document.EndUpdate()

            document.AppendText("Line 1" & vbLf & "Line 2" & vbLf & "Line 3")
            ' Convert all paragraphs to list items.
            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0)
            document.EndUpdate()
            '            #End Region ' #CreateBulletedList
        End Sub


        Private Shared Sub CreateNumberedList(ByVal server As RichEditDocumentServer)
            '            #Region "#CreateNumberedList
            Dim document As Document = server.Document
            document.BeginUpdate()
            Dim abstractListNumberingRoman As AbstractNumberingList = document.AbstractNumberingLists.Add()
            abstractListNumberingRoman.NumberingType = NumberingType.Simple
            Dim level As ListLevel = abstractListNumberingRoman.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.UpperRoman
            level.DisplayFormatString = "{0}."
            Dim numberingList As NumberingList = document.NumberingLists.Add(0)
            document.EndUpdate()
            document.AppendText("Line 1" & vbLf & "Line 2" & vbLf & "Line 3")
            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0)
            document.EndUpdate()
            '            #End Region ' #CreateNumberedList
        End Sub

        Private Shared Sub CreateMultilevelList(ByVal server As RichEditDocumentServer)
            ' #Region "#CreateMultilevelList 
            Dim document As Document = server.Document
            document.BeginUpdate()
            ' Define an abstract list that is the pattern for lists used in the document.
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()
            list.NumberingType = NumberingType.MultiLevel

            ' Specify parameters for each list level.

            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.Decimal
            level.DisplayFormatString = "{0}"

            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 300
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 150
            level.Start = 1
            level.NumberingFormat = NumberingFormat.DecimalEnclosedParenthses
            level.DisplayFormatString = "{0}→{1}"

            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 450
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 220
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerRoman
            level.DisplayFormatString = "{0}→{1}→{2}"

            ' Create a list for use in the document. It is based on a previously defined abstract list with ID = 0.
            document.NumberingLists.Add(0)
            document.EndUpdate()

            document.AppendText("Line one" & vbLf & "Line two" & vbLf & "Line three")
            ' Convert all paragraphs to list items of level 0.
            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs
            For Each pgf As Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = pgf.Index
            Next pgf
            ' Specify a different level for a certain paragraph.
            document.Paragraphs(1).ListLevel = 1
            document.EndUpdate()
            ' #End Region '#CreateMultilevelList
        End Sub

    End Class
End Namespace


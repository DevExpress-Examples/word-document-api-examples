using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ListsActions
    {
        public static Action<RichEditDocumentServer> CreateBulletedListAction = CreateBulletedList;
        public static Action<RichEditDocumentServer> CreateNumberedListAction = CreateNumberedList;
        public static Action<RichEditDocumentServer> CreateMultilevelListAction = CreateMultilevelList;
        static void CreateBulletedList(RichEditDocumentServer wordProcessor)
        {
            #region #CreateBulletedList
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;            

            // Start to edit the document.
            document.BeginUpdate();

            // Create a new list pattern object.
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            // Specify the list type.
            list.NumberingType = NumberingType.Bullet;
            
            // Access the first list level.
            ListLevel level = list.Levels[0];

            // Specify the left indent of the level's paragraph.
            level.ParagraphProperties.LeftIndent = 100;

            // Specify the format of bullets.
            // Without this step, the list is considered as numbered.
            level.DisplayFormatString = "\u00B7";
            level.CharacterProperties.FontName = "Symbol";

            // Create a new list based on the specified pattern.
            NumberingList bulletedList = document.NumberingLists.Add(0);

            // Access the collection of paragraphs.
            ParagraphCollection paragraphs = document.Paragraphs;
            
            // Apply the numbering list format to the document paragraphs.
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0);

            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #CreateBulletedList
        }

        static void CreateNumberedList(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNumberedList
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create a new list pattern object.
            AbstractNumberingList abstractListNumberingRoman = document.AbstractNumberingLists.Add();

            // Specify the list type.
            abstractListNumberingRoman.NumberingType = NumberingType.Simple;

            // Specify properties of the first list level.
            ListLevel level = abstractListNumberingRoman.Levels[0];
            level.ParagraphProperties.LeftIndent = 150;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 75;
            level.Start = 1;

            // Specify the numbering style for the list level.
            level.NumberingFormat = NumberingFormat.LowerRoman;
            level.DisplayFormatString = "{0}.";

            // Create a new list based on the specified pattern.
            NumberingList numberingList = document.NumberingLists.Add(0);

            // Finalize to edit the document.
            document.EndUpdate();

            // Start to edit the document.
            document.BeginUpdate();

            // Access the collection of paragraphs.
            ParagraphCollection paragraphs = document.Paragraphs;

            // Apply the numbering list format to the document paragraphs.
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0);
            
            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #CreateNumberedList
        }

        static void CreateMultilevelList(RichEditDocumentServer wordProcessor)
        {
            #region #CreateMultilevelList
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//List.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            // Start to edit the document.
            document.BeginUpdate();

            // Create a new list pattern object.
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            // Specify the list type.
            list.NumberingType = NumberingType.MultiLevel;

            // Specify parameters for the first list level.
            ListLevel level = list.Levels[0];
            level.ParagraphProperties.LeftIndent = 105;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 55;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.UpperRoman;
            level.DisplayFormatString = "{0}";

            // Specify parameters for the second list level.
            level = list.Levels[1];
            level.ParagraphProperties.LeftIndent = 125;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 65;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.LowerRoman;
            level.DisplayFormatString = "{1})";

            // Specify parameters for the third list level.
            level = list.Levels[2];
            level.ParagraphProperties.LeftIndent = 145;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 75;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.LowerLetter;
            level.DisplayFormatString = "{2}.";

            // Create a new list based on the specified pattern.
            document.NumberingLists.Add(0);

            // Finalize to edit the document.
            document.EndUpdate();

            // Start to edit the document.
            document.BeginUpdate();
            
            // Convert all paragraphs to list items.
            ParagraphCollection paragraphs = document.Paragraphs;
            foreach (Paragraph pgf in paragraphs)
            {
                pgf.ListIndex = 0;
                pgf.ListLevel = pgf.Index;
            }

            // Finalize to edit the document.
            document.EndUpdate();
            #endregion #CreateMultilevelList
        }
    }
}


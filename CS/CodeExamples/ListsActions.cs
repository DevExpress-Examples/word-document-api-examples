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

        static void CreateBulletedList(RichEditDocumentServer server)
        {
            #region #CreateBulletedList
            Document document = server.Document;
            document.LoadDocument("Documents//List.docx");
            document.BeginUpdate();

            // Create a new list pattern object
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            //Specify the list's type
            list.NumberingType = NumberingType.Bullet;
            ListLevel level = list.Levels[0];
            level.ParagraphProperties.LeftIndent = 100;

            //Specify the bullets' format
            //Without this step, the list is considered as numbered
            level.DisplayFormatString = "\u00B7";
            level.CharacterProperties.FontName = "Symbol";

            //Create a new list based on the specific pattern
            NumberingList bulletedList = document.NumberingLists.Add(0);

            // Add paragraphs to the list
            ParagraphCollection paragraphs = document.Paragraphs;
            paragraphs.AddParagraphsToList(document.Range, bulletedList, 0);

            document.EndUpdate();
            #endregion #CreateBulletedList
        }

        static void CreateNumberedList(RichEditDocumentServer server)
        {
            #region #CreateNumberedList
            Document document = server.Document;
            document.LoadDocument("Documents//List.docx");
            document.BeginUpdate();

            //Create a new pattern object
            AbstractNumberingList abstractListNumberingRoman = document.AbstractNumberingLists.Add();

            //Specify the list's type
            abstractListNumberingRoman.NumberingType = NumberingType.Simple;

            //Define the first level's properties
            ListLevel level = abstractListNumberingRoman.Levels[0];
            level.ParagraphProperties.LeftIndent = 150;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 75;
            level.Start = 1;

            //Specify the roman format
            level.NumberingFormat = NumberingFormat.LowerRoman;
            level.DisplayFormatString = "{0}.";

            //Create a new list based on the specific pattern
            NumberingList numberingList = document.NumberingLists.Add(0);

            document.EndUpdate();

            document.BeginUpdate();
            ParagraphCollection paragraphs = document.Paragraphs;
            //Add paragraphs to the list
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0);
            document.EndUpdate();
            #endregion #CreateNumberedList
        }

        static void CreateMultilevelList(RichEditDocumentServer server)
        {
            #region #CreateMultilevelList
            Document document = server.Document;
            document.LoadDocument("Documents//List.docx");
            document.BeginUpdate();

            //Create a new pattern object
            AbstractNumberingList list = document.AbstractNumberingLists.Add();

            //Specify the list's type
            list.NumberingType = NumberingType.MultiLevel;

            //Specify parameters for each list level
            ListLevel level = list.Levels[0];
            level.ParagraphProperties.LeftIndent = 105;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 55;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.UpperRoman;
            level.DisplayFormatString = "{0}";

            level = list.Levels[1];
            level.ParagraphProperties.LeftIndent = 125;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 65;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.LowerRoman;
            level.DisplayFormatString = "{1})";

            level = list.Levels[2];
            level.ParagraphProperties.LeftIndent = 145;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 75;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.LowerLetter;
            level.DisplayFormatString = "{2}.";

            //Create a new list object based on the specified pattern
            document.NumberingLists.Add(0);
            document.EndUpdate();


            //Convert all paragraphs to list items
            document.BeginUpdate();
            ParagraphCollection paragraphs = document.Paragraphs;

            foreach (Paragraph pgf in paragraphs)
            {
                pgf.ListIndex = 0;
                pgf.ListLevel = pgf.Index;
            }

            document.EndUpdate();
            #endregion #CreateMultilevelList
        }
    }
}


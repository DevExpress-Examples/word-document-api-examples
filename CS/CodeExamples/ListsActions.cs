using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    class ListsActions    {
        
        static void CreateBulletedList(RichEditDocumentServer server)
        {
            #region #CreateBulletedList
            Document document = server.Document;
            document.BeginUpdate();
            // Define an abstract list that is the pattern for lists used in the document.
            AbstractNumberingList list = document.AbstractNumberingLists.Add();
            list.NumberingType = NumberingType.Bullet;

            // Specify parameters for each list level.

            ListLevel level = list.Levels[0];
            level.ParagraphProperties.LeftIndent = 100;
            level.CharacterProperties.FontName = "Symbol";
            level.DisplayFormatString = new string('\u00B7', 1);

            level = list.Levels[1];
            level.ParagraphProperties.LeftIndent = 250;
            level.CharacterProperties.FontName = "Symbol";
            level.DisplayFormatString = new string('\u006F', 1);

            level = list.Levels[2];
            level.ParagraphProperties.LeftIndent = 450;
            level.CharacterProperties.FontName = "Symbol";
            level.DisplayFormatString = new string('\u00B7', 1);

            // Create a list for use in the document. It is based on a previously defined abstract list with ID = 0.
            document.NumberingLists.Add(0);
            document.EndUpdate();

            document.AppendText("Line 1\nLine 2\nLine 3");
            // Convert all paragraphs to list items.
            document.BeginUpdate();
            ParagraphCollection paragraphs = document.Paragraphs;
            foreach (Paragraph pgf in paragraphs)
            {
                pgf.ListIndex = 0;
                pgf.ListLevel = 1;
            }
            document.EndUpdate();
            #endregion #CreateBulletedList
        }


        static void CreateNumberedList(RichEditDocumentServer server)
        {
            #region #CreateNumberedList
            Document document = server.Document;
            document.BeginUpdate();
            // Define an abstract list that is the pattern for lists used in the document.
            AbstractNumberingList list = document.AbstractNumberingLists.Add();
            list.NumberingType = NumberingType.MultiLevel;

            // Specify parameters for each list level.

            ListLevel level = list.Levels[0];
            level.ParagraphProperties.LeftIndent = 150;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 75;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.Decimal;
            level.DisplayFormatString = "{0}";

            level = list.Levels[1];
            level.ParagraphProperties.LeftIndent = 300;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 150;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.DecimalEnclosedParenthses;
            level.DisplayFormatString = "{0}→{1}";

            level = list.Levels[2];
            level.ParagraphProperties.LeftIndent = 450;
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
            level.ParagraphProperties.FirstLineIndent = 220;
            level.Start = 1;
            level.NumberingFormat = NumberingFormat.LowerRoman;
            level.DisplayFormatString = "{0}→{1}→{2}";

            // Create a list for use in the document. It is based on a previously defined abstract list with ID = 0.
            document.NumberingLists.Add(0);
            document.EndUpdate();

            document.AppendText("Line one\nLine two\nLine three\nLine four");
            // Convert all paragraphs to list items of level 0.
            document.BeginUpdate();
            ParagraphCollection paragraphs = document.Paragraphs;
            foreach (Paragraph pgf in paragraphs)
            {
                pgf.ListIndex = 0;
                pgf.ListLevel = 0;
            }
            // Specify a different level for a certain paragraph.
            document.Paragraphs[1].ListLevel = 1;
            document.EndUpdate();
            #endregion #CreateNumberedList
        }

    }
}


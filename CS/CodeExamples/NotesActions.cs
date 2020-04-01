using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class NotesActions
    {

        static void InsertFootnotes(RichEditDocumentServer wordProcessor)
        {
            #region #InsertFootnotes
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;
            
            //Insert a footnote at the end of the 6th paragraph:
            DocumentPosition footnotePosition = document.CreatePosition(document.Paragraphs[5].Range.End.ToInt() - 1);
            document.Footnotes.Insert(footnotePosition);

            //Insert a footnote at the end of the 8th paragraph with a custom mark:
            DocumentPosition footnoteWithCustomMarkPosition = document.CreatePosition(document.Paragraphs[7].Range.End.ToInt() - 1);
            document.Footnotes.Insert(footnoteWithCustomMarkPosition, "\u00BA");
            #endregion #InsertFootnotes 
        }


        static void InsertEndnotes(RichEditDocumentServer wordProcessor)
        {
            #region #InsertEndnotes
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;
            
            //Insert an endnote at the end of the last paragraph:
            DocumentPosition endnotePosition = document.CreatePosition(document.Paragraphs[document.Paragraphs.Count - 1].Range.End.ToInt() - 1);
            document.Endnotes.Insert(endnotePosition);

            //Insert an endnote at the end of the second last paragraph with a custom mark:
            DocumentPosition endnoteWithCustomMarkPosition = document.CreatePosition(document.Paragraphs[document.Paragraphs.Count - 2].Range.End.ToInt() - 1);
            document.Endnotes.Insert(endnoteWithCustomMarkPosition, "\u0060");
            #endregion #InsertEndnotes
        }

        static void EditFootnote(RichEditDocumentServer wordProcessor)
        {
            #region #EditFootnote
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;

            //Access the first footnote's content:
            SubDocument footnote = document.Footnotes[0].BeginUpdate();
            
            //Exclude the reference mark and the space after it from the range to be edited:
            DocumentRange noteTextRange = footnote.CreateRange(footnote.Range.Start.ToInt() + 2, footnote.Range.Length
                - 2);
            
            //Clear the range:
            footnote.Delete(noteTextRange);
            
            //Append a new text:
            footnote.AppendText("the text is removed");
            
            //Finalize the update:
            document.Footnotes[0].EndUpdate(footnote);
            #endregion #EditFootnote
        }

        static void EditEndnote(RichEditDocumentServer wordProcessor)
        {
            #region #EditEndnote
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;

            //Access the first endnote's content:
            SubDocument endnote = document.Endnotes[0].BeginUpdate();

            //Exclude the reference mark and the space after it from the range to be edited:
            DocumentRange noteTextRange = endnote.CreateRange(endnote.Range.Start.ToInt() + 2, endnote.Range.Length
                - 2);

            //Access the range's character properties:
            CharacterProperties characterProperties = endnote.BeginUpdateCharacters(noteTextRange);
            
            characterProperties.ForeColor = System.Drawing.Color.Red;
            characterProperties.Italic = true;
            
            //Finalize the character options update:
            endnote.EndUpdateCharacters(characterProperties);
            
            //Finalize the endnote update:
            document.Endnotes[0].EndUpdate(endnote);
            #endregion #EditEndnote
        }

        static void EditSeparator(RichEditDocumentServer wordProcessor)
        {
            #region #EditSeparator
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;
            
            //Check whether the footnotes already have a separator:
            if (document.Footnotes.HasSeparator(NoteSeparatorType.Separator))
            {
                //Initiate the update session:
                SubDocument noteSeparator = document.Footnotes.BeginUpdateSeparator(NoteSeparatorType.Separator);
                
                //Clear the separator range:
                noteSeparator.Delete(noteSeparator.Range);
                
                //Append a new text:
                noteSeparator.AppendText("***");
                
                //Finalize the update:
                document.Footnotes.EndUpdateSeparator(noteSeparator);
            }
            #endregion #EditSeparator
        }
        static void RemoveNotes(RichEditDocumentServer wordProcessor)
        {
            #region #RemoveNotes
            wordProcessor.LoadDocument("Documents//Grimm.docx");
            Document document = wordProcessor.Document;
            
            //Remove first footnote:
            document.Footnotes.RemoveAt(0);
            

            //Remove all custom endnotes:
            for (int i = document.Endnotes.Count - 1; i >= 0; i--)
            {
                if (document.Endnotes[i].IsCustom)
                    document.Endnotes.Remove(document.Endnotes[i]);
            }

            #endregion #RemoveNotes
        }
    }
}

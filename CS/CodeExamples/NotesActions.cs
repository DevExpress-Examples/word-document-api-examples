using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public static class NotesActions
    {
        public static Action<RichEditDocumentServer> InsertFootnotesAction = InsertFootnotes;
        public static Action<RichEditDocumentServer> InsertEndnotesAction = InsertEndnotes;
        public static Action<RichEditDocumentServer> EditFootnoteAction = EditFootnote;
        public static Action<RichEditDocumentServer> EditEndnoteAction = EditEndnote;
        public static Action<RichEditDocumentServer> EditSeparatorAction = EditSeparator;
        public static Action<RichEditDocumentServer> RemoveNotesAction = RemoveNotes;
        static void InsertFootnotes(RichEditDocumentServer wordProcessor)
        {
            #region #InsertFootnotes
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Paragraphs.Count > 5)
            {
                // Insert a footnote at the end of the sixth paragraph.
                DocumentPosition footnotePosition = document.CreatePosition(document.Paragraphs[5].Range.End.ToInt() - 1);
                document.Footnotes.Insert(footnotePosition);

                // Insert a footnote at the end of the eighth paragraph with a custom mark.
                DocumentPosition footnoteWithCustomMarkPosition = document.CreatePosition(document.Paragraphs[7].Range.End.ToInt() - 1);
                document.Footnotes.Insert(footnoteWithCustomMarkPosition, "\u00BA");
            }
            #endregion #InsertFootnotes 
        }


        static void InsertEndnotes(RichEditDocumentServer wordProcessor)
        {
            #region #InsertEndnotes
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;
            
            // Insert an endnote at the end of the last paragraph.
            DocumentPosition endnotePosition = document.CreatePosition(document.Paragraphs[document.Paragraphs.Count - 1].Range.End.ToInt() - 1);
            document.Endnotes.Insert(endnotePosition);

            // Insert an endnote at the end of the second last paragraph with a custom mark.
            DocumentPosition endnoteWithCustomMarkPosition = document.CreatePosition(document.Paragraphs[document.Paragraphs.Count - 2].Range.End.ToInt() - 1);
            document.Endnotes.Insert(endnoteWithCustomMarkPosition, "\u0060");
            #endregion #InsertEndnotes
        }

        static void EditFootnote(RichEditDocumentServer wordProcessor)
        {
            #region #EditFootnote
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the first footnote content.
            SubDocument footnote = document.Footnotes[0].BeginUpdate();
            
            // Exclude the reference mark and the space after it from the range that is edited.
            DocumentRange noteTextRange = footnote.CreateRange(footnote.Range.Start.ToInt() + 2, 
                footnote.Range.Length - 2);
            
            // Clear the range.
            footnote.Delete(noteTextRange);
            
            // Change the footnote text.
            footnote.AppendText("the text is removed");
            
            // Finalize to update the endnote.
            document.Footnotes[0].EndUpdate(footnote);
            #endregion #EditFootnote
        }

        static void EditEndnote(RichEditDocumentServer wordProcessor)
        {
            #region #EditEndnote
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;

            // Access the first endnote content.
            SubDocument endnote = document.Endnotes[0].BeginUpdate();

            // Exclude the reference mark and the space after it from the range that is edited.
            DocumentRange noteTextRange = endnote.CreateRange(endnote.Range.Start.ToInt() + 2, endnote.Range.Length
                - 2);

            // Access the endnote's character formatting.
            CharacterProperties characterProperties = endnote.BeginUpdateCharacters(noteTextRange);

            // Specify the endnote's character formatting options.
            characterProperties.ForeColor = System.Drawing.Color.Red;
            characterProperties.Italic = true;
            
            // Finalize to update character formatting.
            endnote.EndUpdateCharacters(characterProperties);
            
            // Finalize to update the endnote.
            document.Endnotes[0].EndUpdate(endnote);
            #endregion #EditEndnote
        }

        static void EditSeparator(RichEditDocumentServer wordProcessor)
        {
            #region #EditSeparator
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;
            
            // Check whether footnotes already have a separator.
            if (document.Footnotes.HasSeparator(NoteSeparatorType.Separator))
            {
                // Access the footnote separator.
                SubDocument noteSeparator = document.Footnotes.BeginUpdateSeparator(NoteSeparatorType.Separator);
                
                // Clear the separator range.
                noteSeparator.Delete(noteSeparator.Range);
                
                // Change the footnote separator.
                noteSeparator.AppendText("***");
                
                // Finalize to update the footnote separator.
                document.Footnotes.EndUpdateSeparator(noteSeparator);
            }
            #endregion #EditSeparator
        }
        static void RemoveNotes(RichEditDocumentServer wordProcessor)
        {
            #region #RemoveNotes
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents//Grimm.docx");

            // Access a document.
            Document document = wordProcessor.Document;
            
            if (document.Footnotes.Count > 0)
                // Remove the first footnote.
                document.Footnotes.RemoveAt(0);
            

            // Remove all custom endnotes.
            for (int i = document.Endnotes.Count - 1; i >= 0; i--)
            {
                if (document.Endnotes[i].IsCustom)
                    document.Endnotes.Remove(document.Endnotes[i]);
            }

            #endregion #RemoveNotes
        }
    }
}

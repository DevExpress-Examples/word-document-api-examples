using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;
using System.Diagnostics;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public class CommentsActions
    {
        public static Action<RichEditDocumentServer> CreateCommentAction = CreateComment;
        public static Action<RichEditDocumentServer> CreateNestedCommentAction = CreateNestedComment;
        public static Action<RichEditDocumentServer> DeleteCommentAction = DeleteComment;
        public static Action<RichEditDocumentServer> EditCommentPropertiesAction = EditCommentProperties;
        public static Action<RichEditDocumentServer> EditCommentContentAction = EditCommentContent;

        static void CreateComment(RichEditDocumentServer wordProcessor)
        {
            #region #CreateComment
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Paragraphs.Count > 2)
            {
                // Access the range of the third paragraph.
                DocumentRange docRange = document.Paragraphs[2].Range;

                // Specify the comment's author name.
                string commentAuthor = "Johnson Alphonso D";

                // Create a comment.
                document.Comments.Create(docRange, commentAuthor, DateTime.Now);
            }
            #endregion #CreateComment
        }

        static void CreateNestedComment(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNestedComment
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Comments.Count > 1)
            {
                // Find text ranges matched the string in the document range
                // to which the parent comment relates.
                DocumentRange[] resRanges = document.FindAll("trump", SearchOptions.None, document.Comments[1].Range);
                if (resRanges.Length > 0)
                {
                    // Create a new comment nested in the parent comment.
                    Comment newComment = document.Comments.Create("Vicars Anny", document.Comments[1]);
                    newComment.Date = DateTime.Now;
                }
            }
            #endregion #CreateNestedComment
        }

        static void DeleteComment(RichEditDocumentServer wordProcessor)
        {
            #region #DeleteComment
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            if (document.Comments.Count > 0)
            {
                // Delete the first comment.
                document.Comments.Remove(document.Comments[0]);
            }
            #endregion #DeleteComment
        }

        static void EditCommentProperties(RichEditDocumentServer wordProcessor)
        {
            #region #EditCommentProperties
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            int commentCount = document.Comments.Count;
            if (commentCount > 0)
            {
                // Start to edit the document.
                document.BeginUpdate();

                // Access a comment and edit its properties.
                Comment comment = document.Comments[document.Comments.Count - 1];
                comment.Name = "New Name";
                comment.Date = DateTime.Now;
                comment.Author = "New Author";
                
                // Finalize to edit the document.
                document.EndUpdate();
            }
            #endregion #EditCommentProperties
        }

        static void EditCommentContent(RichEditDocumentServer wordProcessor)
        {
            #region #EditCommentContent
            // Load a document from a file.
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);

            // Access a document.
            Document document = wordProcessor.Document;

            int commentCount = document.Comments.Count;
            if (commentCount > 0)
            {
                // Access a comment.
                Comment comment = document.Comments[document.Comments.Count - 1];
                if (comment != null)
                {
                    // Start to edit the comment.
                    SubDocument commentDocument = comment.BeginUpdate();

                    // Insert a text to the comment.
                    commentDocument.InsertText(commentDocument.CreatePosition(0), "some text");

                    // Insert a table to the comment.
                    commentDocument.Tables.Create(commentDocument.CreatePosition(9), 5, 4);

                    // Finalize to edit the comment.
                    comment.EndUpdate(commentDocument);
                }
            }
            #endregion #EditCommentContent
        }


    }
}

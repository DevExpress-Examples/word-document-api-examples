using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

namespace RichEditDocumentServerAPIExample.CodeExamples
{
    public class CommentsActions
    {
         void CreateComment(RichEditDocumentServer wordProcessor)
        {
            #region #CreateComment
            Document document = wordProcessor.Document;
            wordProcessor.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            DocumentRange docRange = document.Paragraphs[2].Range;
            string commentAuthor = "Johnson Alphonso D";
            document.Comments.Create(docRange, commentAuthor, DateTime.Now);
            #endregion #CreateComment
        }

         void CreateNestedComment(RichEditDocumentServer wordProcessor)
        {
            #region #CreateNestedComment
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            if (document.Comments.Count > 0)
            {
                DocumentRange[] resRanges = document.FindAll("trump", SearchOptions.None, document.Comments[1].Range);
                if (resRanges.Length > 0)
                {
                    Comment newComment = document.Comments.Create("Vicars Anny", document.Comments[1]);
                    newComment.Date = DateTime.Now;
                }
            }
            #endregion #CreateNestedComment
        }

         void DeleteComment(RichEditDocumentServer wordProcessor)
        {
            #region #DeleteComment
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            if (document.Comments.Count > 0)
            {
                document.Comments.Remove(document.Comments[0]);
            }
            #endregion #DeleteComment
        }

         void EditCommentProperties(RichEditDocumentServer wordProcessor)
        {
            #region #EditCommentProperties
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            int commentCount = document.Comments.Count;
            if (commentCount > 0)
            {
                document.BeginUpdate();
                Comment comment = document.Comments[document.Comments.Count - 1];
                comment.Name = "New Name";
                comment.Date = DateTime.Now;
                comment.Author = "New Author";
                document.EndUpdate();
            }
            #endregion #EditCommentProperties
        }

         void EditCommentContent(RichEditDocumentServer wordProcessor)
        {
            #region #EditCommentContent
            Document document = wordProcessor.Document;
            document.LoadDocument("Documents\\Grimm.docx", DocumentFormat.OpenXml);
            int commentCount = document.Comments.Count;
            if (commentCount > 0)
            {
                Comment comment = document.Comments[document.Comments.Count - 1];
                if (comment != null)
                {
                    SubDocument commentDocument = comment.BeginUpdate();
                    commentDocument.InsertText(commentDocument.CreatePosition(0), "some text");
                    commentDocument.Tables.Create(commentDocument.CreatePosition(9), 5, 4);
                    comment.EndUpdate(commentDocument);
                }
            }
            #endregion #EditCommentContent
        }


    }
}

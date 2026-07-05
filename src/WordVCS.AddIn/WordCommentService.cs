using System;
using System.Collections.Generic;
using WordVCS.Core.Models;
using WordVCS.Core.Services;
using Word = Microsoft.Office.Interop.Word;

namespace WordVCS.AddIn
{
    /// <summary>
    /// Extracts live comments from an open Word document via COM.
    /// </summary>
    public class WordCommentService
    {
        private readonly ICommentExtractor _fileExt;

        public WordCommentService(ICommentExtractor fileExt) { _fileExt = fileExt; }

        public List<CommentRecord> ExtractFromDocument(Word.Document doc)
        {
            var r = new List<CommentRecord>();
            if (doc == null) return r;
            try
            {
                foreach (Word.Comment c in doc.Comments)
                {
                    try
                    {
                        var rec = new CommentRecord
                        {
                            Author = c.Author ?? "Unknown",
                            Content = (c.Range?.Text ?? "").Trim(),
                            CreatedAt = c.Date,
                            RangeStart = c.Scope?.Start ?? 0,
                            RangeEnd = c.Scope?.End ?? 0,
                            SelectedText = (c.Scope?.Text ?? "").Trim(),
                            Status = CommentStatus.Pending
                        };
                        rec.Id = _fileExt.GenerateCommentId(
                            rec.Author, rec.Content, rec.SelectedText, rec.RangeStart);
                        r.Add(rec);
                    }
                    catch { }
                }
            }
            catch { }
            return r;
        }
    }
}

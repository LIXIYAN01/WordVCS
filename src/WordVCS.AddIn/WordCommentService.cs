using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using WordVCS.Core.Models;
using WordVCS.Core.Services;

namespace WordVCS.AddIn
{
    /// <summary>
    /// Word 实时批注服务 — 封装 Word Object Model 的 COM 调用。
    /// 提供从当前打开的 Word 文档中提取批注的能力。
    /// </summary>
    public class WordCommentService
    {
        private readonly ICommentExtractor _fileExtractor;

        public WordCommentService(ICommentExtractor fileExtractor)
        {
            _fileExtractor = fileExtractor;
        }

        /// <summary>
        /// 从当前打开的 Word Document 中提取所有批注（Word COM）
        /// </summary>
        public List<CommentRecord> ExtractFromDocument(Word.Document doc)
        {
            var records = new List<CommentRecord>();
            if (doc == null) return records;

            try
            {
                foreach (Word.Comment comment in doc.Comments)
                {
                    try
                    {
                        var record = new CommentRecord
                        {
                            Author = comment.Author ?? "未知作者",
                            Content = comment.Range?.Text?.Trim() ?? "",
                            CreatedAt = comment.Date,
                            RangeStart = comment.Scope?.Start ?? 0,
                            RangeEnd = comment.Scope?.End ?? 0,
                            SelectedText = comment.Scope?.Text?.Trim() ?? "",
                            Status = CommentStatus.Pending
                        };

                        record.Id = _fileExtractor.GenerateCommentId(
                            record.Author, record.Content,
                            record.SelectedText, record.RangeStart);

                        records.Add(record);
                    }
                    catch (Exception)
                    {
                        // Skip malformed comments
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[WordVCS] ExtractFromDocument: " + ex.Message);
            }

            return records;
        }

        /// <summary>
        /// 检查文档是否有新增或删除的批注
        /// </summary>
        public List<CommentChange> DetectCommentChanges(
            Word.Document doc, List<CommentRecord> baseline)
        {
            var current = ExtractFromDocument(doc);
            return _fileExtractor.CompareComments(current, baseline);
        }
    }
}

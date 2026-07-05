using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 基于 Open XML SDK 的批注提取器。
    /// 完全离线，仅从 .docx 文件内部的 word/comments.xml 读取批注。
    /// </summary>
    public class CommentExtractor : ICommentExtractor
    {
        public List<CommentRecord> ExtractFromFile(string filePath)
        {
            var records = new List<CommentRecord>();
            if (!File.Exists(filePath)) return records;

            try
            {
                using (var wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    var commentsPart = wordDoc.MainDocumentPart?.WordprocessingCommentsPart;
                    if (commentsPart?.Comments == null) return records;

                    foreach (var c in commentsPart.Comments
                        .Elements<DocumentFormat.OpenXml.Wordprocessing.Comment>())
                    {
                        try
                        {
                            var author = c.Author?.Value ?? "未知作者";
                            var content = c.InnerText ?? string.Empty;
                            var date = c.Date?.Value ?? DateTime.MinValue;

                            var record = new CommentRecord
                            {
                                Author = author,
                                Content = content.Trim(),
                                CreatedAt = date,
                                Status = CommentStatus.Pending
                            };
                            record.Id = GenerateCommentId(author, content,
                                string.Empty, 0);
                            records.Add(record);
                        }
                        catch { /* skip malformed comment */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordVCS] ExtractFromFile failed: {ex.Message}");
            }
            return records;
        }

        public List<CommentRecord> ExtractFromBytes(byte[] docxBytes,
            string sourceName)
        {
            if (docxBytes == null || docxBytes.Length == 0)
                return new List<CommentRecord>();

            var tmpPath = Path.Combine(Path.GetTempPath(),
                $"wordvcs_extract_{Guid.NewGuid():N}.docx");
            try
            {
                File.WriteAllBytes(tmpPath, docxBytes);
                return ExtractFromFile(tmpPath);
            }
            finally
            {
                try { File.Delete(tmpPath); } catch { }
            }
        }

        public string GenerateCommentId(string author, string content,
            string selectedText, int rangeStart)
        {
            var cleanAuthor = author ?? "Unknown";
            var cleanContent = content?.Trim() ?? string.Empty;
            var cleanSelected = selectedText?.Trim() ?? string.Empty;

            if (cleanContent.Length > 50)
                cleanContent = cleanContent.Substring(0, 50);
            if (cleanSelected.Length > 30)
                cleanSelected = cleanSelected.Substring(0, 30);

            var rawKey =
                $"{cleanAuthor}_{cleanContent}_{cleanSelected}_{rangeStart / 1000}";

            using (var sha1 = SHA1.Create())
            {
                var hash = sha1.ComputeHash(Encoding.UTF8.GetBytes(rawKey));
                var sb = new StringBuilder();
                for (int i = 0; i < Math.Min(hash.Length, 6); i++)
                    sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }

        public List<CommentChange> CompareComments(
            List<CommentRecord> current, List<CommentRecord> baseline)
        {
            var changes = new List<CommentChange>();
            var baselineMap = baseline.ToDictionary(b => b.Id, b => b);

            foreach (var cc in current)
            {
                if (baselineMap.TryGetValue(cc.Id, out var bc))
                {
                    if (cc.Status != bc.Status)
                    {
                        changes.Add(new CommentChange
                        {
                            CommentId = cc.Id,
                            ChangeType = "Modified",
                            OldRecord = bc,
                            NewRecord = cc
                        });
                    }
                    baselineMap.Remove(cc.Id);
                }
                else
                {
                    changes.Add(new CommentChange
                    {
                        CommentId = cc.Id,
                        ChangeType = "Added",
                        NewRecord = cc
                    });
                }
            }

            // Remaining baseline comments were removed
            foreach (var kv in baselineMap)
            {
                changes.Add(new CommentChange
                {
                    CommentId = kv.Key,
                    ChangeType = "Removed",
                    OldRecord = kv.Value
                });
            }

            return changes;
        }
    }
}

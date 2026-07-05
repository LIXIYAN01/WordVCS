using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using DocumentFormat.OpenXml.Packaging;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 差异服务实现，基于 Open XML 文本提取与 DiffPlex 算法引擎
    /// </summary>
    public class DiffService : IDiffService
    {
        public Task<string> ExtractTextFromDocx(string filePath)
        {
            return Task.Run(() =>
            {
                if (!File.Exists(filePath)) return string.Empty;
                try
                {
                    var sb = new StringBuilder();
                    using (var wordDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        var body = wordDoc.MainDocumentPart?.Document?.Body;
                        if (body == null) return string.Empty;

                        foreach (var para in body
                            .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                        {
                            sb.AppendLine(para.InnerText);
                        }
                    }
                    return sb.ToString();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordVCS] ExtractText failed: {ex.Message}");
                    return string.Empty;
                }
            });
        }

        public Task<DiffResult> CompareText(string oldText, string newText)
        {
            return Task.Run(() =>
            {
                var result = new DiffResult();
                try
                {
                    var differ = new Differ();
                    var builder = new InlineDiffBuilder(differ);
                    var diffModel = builder.BuildDiffModel(
                        oldText ?? string.Empty, newText ?? string.Empty);

                    int added = 0, removed = 0, modified = 0;
                    int lineNum = 1;

                    foreach (var line in diffModel.Lines)
                    {
                        var block = new DiffBlock
                        {
                            Text = line.Text,
                            OldLineStart = lineNum,
                            NewLineStart = lineNum,
                            OldLineCount = 1,
                            NewLineCount = 1
                        };

                        switch (line.Type)
                        {
                            case ChangeType.Inserted:
                                block.Type = DiffType.Added;
                                added++;
                                break;
                            case ChangeType.Deleted:
                                block.Type = DiffType.Removed;
                                removed++;
                                break;
                            case ChangeType.Modified:
                                block.Type = DiffType.Modified;
                                modified++;
                                break;
                            default:
                                block.Type = DiffType.Unchanged;
                                break;
                        }

                        result.Blocks.Add(block);
                        lineNum++;
                    }

                    result.Stats = new DiffStats
                    {
                        LinesAdded = added,
                        LinesRemoved = removed,
                        LinesModified = modified,
                        WordsAdded = added * 5,
                        WordsRemoved = removed * 5
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordVCS] CompareText failed: {ex.Message}");
                }
                return result;
            });
        }

        public async Task<DiffResult> CompareDocx(string oldFilePath,
            string newFilePath)
        {
            var oldText = await ExtractTextFromDocx(oldFilePath);
            var newText = await ExtractTextFromDocx(newFilePath);
            return await CompareText(oldText, newText);
        }
    }
}

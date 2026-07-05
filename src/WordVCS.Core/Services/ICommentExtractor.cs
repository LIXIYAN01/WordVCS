using System.Collections.Generic;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 批注提取服务接口 — 仅使用 Open XML SDK 从 .docx 文件中提取批注，
    /// 不依赖 Microsoft.Office.Interop.Word（COM）。
    /// 实时 Word 文档的批注提取由 VSTO AddIn 层的 WordCommentService 负责。
    /// </summary>
    public interface ICommentExtractor
    {
        /// <summary>
        /// 从 .docx 文件提取所有批注（通过 Open XML SDK，无需启动 Word）
        /// </summary>
        List<CommentRecord> ExtractFromFile(string filePath);

        /// <summary>
        /// 从 .docx 原始字节流提取批注（用于恢复历史版本）
        /// </summary>
        List<CommentRecord> ExtractFromBytes(byte[] docxBytes, string sourceName);

        /// <summary>
        /// 生成批注的全局唯一哈希标识，用于跨版本追踪
        /// </summary>
        string GenerateCommentId(string author, string content,
            string selectedText, int rangeStart);

        /// <summary>
        /// 对比两个批注列表，标记新增/已解决/变更
        /// </summary>
        List<CommentChange> CompareComments(List<CommentRecord> current,
            List<CommentRecord> baseline);
    }
}

using System;
using System.Collections.Generic;

namespace WordVCS.Core.Models
{
    /// <summary>
    /// 表示一次版本提交记录（对应 Git 中的 Commit）
    /// </summary>
    public class VersionRecord
    {
        public string Sha { get; set; }
        public string ShortSha => !string.IsNullOrEmpty(Sha) && Sha.Length >= 7 ? Sha.Substring(0, 7) : Sha;
        public string Author { get; set; }
        public string Email { get; set; }
        public string Message { get; set; }
        public string ShortMessage
        {
            get
            {
                if (string.IsNullOrEmpty(Message)) return string.Empty;
                var lines = Message.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                return lines.Length > 0 ? lines[0] : string.Empty;
            }
        }
        public DateTime CommittedAt { get; set; }
        public string ParentSha { get; set; }
        public string Branch { get; set; }
        public List<string> Tags { get; set; } = new List<string>();
        public List<CommentRecord> AddressedComments { get; set; } = new List<CommentRecord>();
    }

    /// <summary>
    /// 表示一个批注记录
    /// </summary>
    public class CommentRecord
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public string Content { get; set; }
        public DateTime CreatedAt { get; set; }
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
        public string SelectedText { get; set; }
        public CommentStatus Status { get; set; } = CommentStatus.Pending;
        public string ResolvedInCommit { get; set; }
        public DateTime LastUpdatedAt { get; set; }
        public string FirstSeenSha { get; set; }
    }

    /// <summary>
    /// 批注处理状态
    /// </summary>
    public enum CommentStatus
    {
        Pending,
        Addressed,
        Dismissed,
        Replied
    }

    /// <summary>
    /// 批注与提交的关联映射关系
    /// </summary>
    public class CommentMapping
    {
        public int Id { get; set; }
        public string CommentId { get; set; }
        public string CommitSha { get; set; }
        public string Action { get; set; }
        public DateTime MappedAt { get; set; }
    }

    /// <summary>
    /// 表示批注在文档生命周期中的全景图
    /// </summary>
    public class CommentLifecycle
    {
        public CommentRecord Comment { get; set; }
        public VersionRecord CreatedInVersion { get; set; }
        public VersionRecord ResolvedInVersion { get; set; }
        public List<VersionRecord> RelatedVersions { get; set; } = new List<VersionRecord>();
    }

    /// <summary>
    /// 差异对比类型
    /// </summary>
    public enum DiffType
    {
        Unchanged,
        Added,
        Removed,
        Modified
    }

    /// <summary>
    /// 差异对比的数据分块
    /// </summary>
    public class DiffBlock
    {
        public DiffType Type { get; set; }
        public string Text { get; set; }
        public string OldText { get; set; }
        public string NewText { get; set; }
        public int OldLineStart { get; set; }
        public int OldLineCount { get; set; }
        public int NewLineStart { get; set; }
        public int NewLineCount { get; set; }
    }

    /// <summary>
    /// 差异对比汇总结果
    /// </summary>
    public class DiffResult
    {
        public List<DiffBlock> Blocks { get; set; } = new List<DiffBlock>();
        public DiffStats Stats { get; set; } = new DiffStats();
    }

    /// <summary>
    /// 差异统计指标
    /// </summary>
    public class DiffStats
    {
        public int LinesAdded { get; set; }
        public int LinesRemoved { get; set; }
        public int LinesModified { get; set; }
        public int WordsAdded { get; set; }
        public int WordsRemoved { get; set; }
    }

    /// <summary>
    /// 分支信息
    /// </summary>
    public class BranchInfo
    {
        public string Name { get; set; }
        public bool IsCurrent { get; set; }
        public bool IsRemote { get; set; }
        public string HeadCommitSha { get; set; }
    }

    /// <summary>
    /// 仓库初始化及操作配置
    /// </summary>
    public class RepositoryConfig
    {
        public string AuthorName { get; set; }
        public string AuthorEmail { get; set; }
        public string DefaultBranch { get; set; } = "draft";
        public bool AutoSaveEnabled { get; set; } = true;
        public int AutoSaveIntervalMinutes { get; set; } = 15;
    }

    /// <summary>
    /// 全局软件配置
    /// </summary>
    public class UserSettings
    {
        public string Language { get; set; } = "zh-CN";
        public string Theme { get; set; } = "Light";
        public double TaskPaneWidth { get; set; } = 360;
        public string DefaultAuthorName { get; set; }
        public string DefaultAuthorEmail { get; set; }
    }

    /// <summary>
    /// 仓库状态信息
    /// </summary>
    public class RepositoryStatus
    {
        public bool IsRepository { get; set; }
        public string CurrentBranch { get; set; }
        public string HeadCommitSha { get; set; }
        public string HeadCommitMessage { get; set; }
        public bool HasUncommittedChanges { get; set; }
        public int UnresolvedCommentCount { get; set; }
        public int TotalCommentCount { get; set; }
        public string LastAutoSave { get; set; }
    }

    /// <summary>
    /// 批注变更记录
    /// </summary>
    public class CommentChange
    {
        public string CommentId { get; set; }
        public string ChangeType { get; set; } // Added, Resolved, Modified, Removed
        public CommentRecord OldRecord { get; set; }
        public CommentRecord NewRecord { get; set; }
    }
}

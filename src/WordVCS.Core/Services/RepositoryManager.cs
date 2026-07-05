using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// 仓库管理器实现，协调 Git、批注提取、SQLite 映射的核心编排层
    /// </summary>
    public class RepositoryManager : IRepositoryManager
    {
        private readonly IGitService _git;
        private readonly ICommentExtractor _commentExtractor;
        private readonly IMappingService _mapping;

        public RepositoryManager(
            IGitService gitService,
            ICommentExtractor commentExtractor,
            IMappingService mappingService)
        {
            _git = gitService;
            _commentExtractor = commentExtractor;
            _mapping = mappingService;
        }

        public async Task<bool> EnsureRepository(string docPath,
            string authorName, string authorEmail, string defaultBranch)
        {
            if (string.IsNullOrEmpty(docPath)) return false;

            if (!_git.IsRepository(docPath))
            {
                var inited = await _git.InitRepository(
                    docPath, authorName, authorEmail, defaultBranch);
                if (!inited) return false;
            }

            _mapping.InitializeDatabase(docPath);
            return true;
        }

        public async Task<string> ExecuteCommit(string docPath, string message,
            string authorName, string authorEmail,
            List<CommentRecord> addressedComments)
        {
            try
            {
                await EnsureRepository(docPath, authorName, authorEmail, "draft");

                // 1. Extract current comments from file for snapshot
                var currentComments = _commentExtractor.ExtractFromFile(docPath);

                // 2. Merge addressed status into comments
                if (addressedComments != null && addressedComments.Count > 0)
                {
                    var addressedIds = new HashSet<string>(
                        addressedComments.Select(c => c.Id));
                    foreach (var c in currentComments)
                    {
                        if (addressedIds.Contains(c.Id))
                        {
                            c.Status = CommentStatus.Addressed;
                        }
                    }
                }

                // 3. Git commit
                var commitSha = await _git.Commit(
                    docPath, message, authorName, authorEmail);

                if (!string.IsNullOrEmpty(commitSha) &&
                    addressedComments != null && addressedComments.Count > 0)
                {
                    // 4. Persist mapping
                    await _mapping.SaveMapping(docPath, commitSha, addressedComments);
                }

                return commitSha;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordVCS] ExecuteCommit failed: {ex.Message}");
                return string.Empty;
            }
        }

        public async Task<string> TriggerAutoSave(string docPath,
            string authorName, string authorEmail)
        {
            try
            {
                if (!_git.IsRepository(docPath)) return string.Empty;

                var wipMessage =
                    $"[自动保存] WordVCS AutoSave: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                return await _git.Commit(docPath, wipMessage, authorName, authorEmail);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordVCS] AutoSave failed: {ex.Message}");
                return string.Empty;
            }
        }

        public RepositoryStatus GetStatus(string docPath)
        {
            var status = new RepositoryStatus
            {
                IsRepository = _git.IsRepository(docPath)
            };

            if (!status.IsRepository) return status;

            try
            {
                status.CurrentBranch = _git.GetCurrentBranch(docPath);
                status.HasUncommittedChanges = _git.HasUncommittedChanges(docPath);

                var history = _git.GetHistory(docPath, maxCount: 1);
                if (history.Count > 0)
                {
                    status.HeadCommitSha = history[0].ShortSha;
                    status.HeadCommitMessage = history[0].ShortMessage;
                }

                var allComments = _mapping.GetPersistedComments(docPath);
                status.TotalCommentCount = allComments.Count;
                status.UnresolvedCommentCount = allComments
                    .Count(c => c.Status == CommentStatus.Pending);
            }
            catch { }

            return status;
        }
    }
}

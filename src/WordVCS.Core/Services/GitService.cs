using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using LibGit2Sharp;
using WordVCS.Core.Models;

namespace WordVCS.Core.Services
{
    /// <summary>
    /// Git 服务具体实现（基于 LibGit2Sharp）
    /// </summary>
    public class GitService : IGitService
    {
        private static string GetRepositoryPath(string docPath)
        {
            if (string.IsNullOrEmpty(docPath)) return null;
            var dir = Path.GetDirectoryName(docPath);
            // Walk up to find .git directory
            while (!string.IsNullOrEmpty(dir))
            {
                if (Directory.Exists(Path.Combine(dir, ".git")))
                    return dir;
                var parent = Directory.GetParent(dir);
                if (parent == null) break;
                dir = parent.FullName;
            }
            return Path.GetDirectoryName(docPath);
        }

        public bool IsRepository(string docPath)
        {
            var repoDir = GetRepositoryPath(docPath);
            return !string.IsNullOrEmpty(repoDir) && Repository.IsValid(repoDir);
        }

        public Task<bool> InitRepository(string docPath, string authorName,
            string authorEmail, string defaultBranch)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = Path.GetDirectoryName(docPath);
                    if (string.IsNullOrEmpty(repoDir)) return false;

                    Repository.Init(repoDir);

                    using (var repo = new Repository(repoDir))
                    {
                        // Write .gitignore
                        var gitignorePath = Path.Combine(repoDir, ".gitignore");
                        if (!File.Exists(gitignorePath))
                        {
                            File.WriteAllText(gitignorePath,
                                "~$*.docx\r\n~$*.docm\r\n*.tmp\r\n*.wbk\r\n" +
                                "Thumbs.db\r\nDesktop.ini\r\n" +
                                ".wordvcs/snapshots/\r\n.wordvcs/logs/\r\n");
                        }

                        // Write .gitattributes for docx diff
                        var gaPath = Path.Combine(repoDir, ".gitattributes");
                        if (!File.Exists(gaPath))
                        {
                            File.WriteAllText(gaPath,
                                "*.docx diff=wordvcs\r\n*.docm diff=wordvcs\r\n" +
                                "*.pdf binary\r\n*.png binary\r\n*.jpg binary\r\n");
                        }

                        // Initial commit
                        Commands.Stage(repo, "*");
                        var sig = new Signature(authorName, authorEmail, DateTimeOffset.Now);
                        repo.Commit("初始化论文版本控制系统 (WordVCS)", sig, sig);

                        // Handle default branch
                        if (!string.IsNullOrEmpty(defaultBranch) &&
                            defaultBranch != "master" &&
                            repo.Branches["master"] != null)
                        {
                            var branch = repo.CreateBranch(defaultBranch);
                            repo.Refs.UpdateTarget(repo.Refs.Head, branch.CanonicalName);
                            repo.Branches.Remove("master");
                        }
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordVCS] InitRepository failed: {ex.Message}");
                    return false;
                }
            });
        }

        public Task<string> Commit(string docPath, string message,
            string authorName, string authorEmail)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = GetRepositoryPath(docPath);
                    using (var repo = new Repository(repoDir))
                    {
                        Commands.Stage(repo, "*");
                        var status = repo.RetrieveStatus();
                        if (!status.IsDirty) return string.Empty;

                        var sig = new Signature(authorName, authorEmail, DateTimeOffset.Now);
                        var commit = repo.Commit(message, sig, sig);
                        return commit.Sha;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordVCS] Commit failed: {ex.Message}");
                    return string.Empty;
                }
            });
        }

        public List<VersionRecord> GetHistory(string docPath,
            string branchName = null, int maxCount = 100)
        {
            var history = new List<VersionRecord>();
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                if (!IsRepository(docPath)) return history;

                using (var repo = new Repository(repoDir))
                {
                    var currentBranch = repo.Head.FriendlyName;

                    if (!string.IsNullOrEmpty(branchName))
                    {
                        var branch = repo.Branches[branchName];
                        if (branch == null) return history;
                        return BuildHistoryFromLog(branch.Commits, repo, currentBranch, maxCount);
                    }

                    return BuildHistoryFromLog(repo.Commits, repo, currentBranch, maxCount);

                    List<VersionRecord> BuildHistoryFromLog(ICommitLog log, Repository r,
                        string curBranch, int max)
                    {
                        var result = new List<VersionRecord>();
                        var tagMap = r.Tags
                            .GroupBy(t => t.Target.Sha)
                            .ToDictionary(g => g.Key,
                                g => g.Select(t => t.FriendlyName).ToList());

                        int c = 0;
                        foreach (var commit in log)
                        {
                            if (c++ >= max) break;
                            var rec = new VersionRecord
                            {
                                Sha = commit.Sha,
                                Author = commit.Author.Name,
                                Email = commit.Author.Email,
                                Message = commit.Message,
                                CommittedAt = commit.Author.When.DateTime,
                                ParentSha = commit.Parents.FirstOrDefault()?.Sha,
                                Branch = curBranch
                            };
                            if (tagMap.ContainsKey(commit.Sha))
                                rec.Tags = tagMap[commit.Sha];
                            result.Add(rec);
                        }
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordVCS] GetHistory failed: {ex.Message}");
            }
            return history;
        }

        public VersionRecord GetCommitDetail(string docPath, string commitSha)
        {
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                using (var repo = new Repository(repoDir))
                {
                    var commit = repo.Lookup<Commit>(commitSha);
                    if (commit == null) return null;

                    var tagMap = repo.Tags
                        .Where(t => t.Target.Sha == commitSha)
                        .Select(t => t.FriendlyName).ToList();

                    return new VersionRecord
                    {
                        Sha = commit.Sha,
                        Author = commit.Author.Name,
                        Email = commit.Author.Email,
                        Message = commit.Message,
                        CommittedAt = commit.Author.When.DateTime,
                        ParentSha = commit.Parents.FirstOrDefault()?.Sha,
                        Branch = repo.Head.FriendlyName,
                        Tags = tagMap
                    };
                }
            }
            catch { return null; }
        }

        public List<BranchInfo> ListBranches(string docPath)
        {
            var branches = new List<BranchInfo>();
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                using (var repo = new Repository(repoDir))
                {
                    foreach (var b in repo.Branches)
                    {
                        if (b.IsRemote) continue;
                        branches.Add(new BranchInfo
                        {
                            Name = b.FriendlyName,
                            IsCurrent = b.IsCurrentRepositoryHead,
                            IsRemote = b.IsRemote,
                            HeadCommitSha = b.Tip?.Sha
                        });
                    }
                }
            }
            catch { }
            return branches;
        }

        public Task<bool> CreateBranch(string docPath, string branchName)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = GetRepositoryPath(docPath);
                    using (var repo = new Repository(repoDir))
                    {
                        if (repo.Branches[branchName] != null) return false;
                        var newBranch = repo.CreateBranch(branchName);
                        repo.Refs.UpdateTarget(repo.Refs.Head,
                            newBranch.CanonicalName);
                        return true;
                    }
                }
                catch { return false; }
            });
        }

        public Task<bool> SwitchBranch(string docPath, string branchName)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = GetRepositoryPath(docPath);
                    using (var repo = new Repository(repoDir))
                    {
                        var branch = repo.Branches[branchName];
                        if (branch == null) return false;
                        Commands.Checkout(repo, branch);
                        return true;
                    }
                }
                catch { return false; }
            });
        }

        public Task<bool> CreateTag(string docPath, string tagName, string message)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = GetRepositoryPath(docPath);
                    using (var repo = new Repository(repoDir))
                    {
                        repo.ApplyTag(tagName, repo.Head.Tip.Sha,
                            repo.Config.BuildSignature(DateTimeOffset.Now), message);
                        return true;
                    }
                }
                catch { return false; }
            });
        }

        public List<string> ListTags(string docPath)
        {
            var tags = new List<string>();
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                using (var repo = new Repository(repoDir))
                {
                    foreach (var tag in repo.Tags)
                        tags.Add(tag.FriendlyName);
                }
            }
            catch { }
            return tags;
        }

        public Task<bool> RestoreVersion(string docPath, string commitSha,
            string outputPath)
        {
            return Task.Run(() =>
            {
                try
                {
                    var repoDir = GetRepositoryPath(docPath);
                    var fileName = Path.GetFileName(docPath);
                    using (var repo = new Repository(repoDir))
                    {
                        var commit = repo.Lookup<Commit>(commitSha);
                        if (commit == null) return false;
                        var treeEntry = commit[fileName];
                        if (treeEntry == null) return false;
                        var blob = (Blob)treeEntry.Target;
                        using (var contentStream = blob.GetContentStream())
                        using (var fileStream = File.Create(outputPath))
                        {
                            contentStream.CopyTo(fileStream);
                        }
                    }
                    return true;
                }
                catch { return false; }
            });
        }

        public bool HasUncommittedChanges(string docPath)
        {
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                if (!IsRepository(docPath)) return false;
                using (var repo = new Repository(repoDir))
                {
                    return repo.RetrieveStatus().IsDirty;
                }
            }
            catch { return false; }
        }

        public string GetCurrentBranch(string docPath)
        {
            try
            {
                var repoDir = GetRepositoryPath(docPath);
                using (var repo = new Repository(repoDir))
                {
                    return repo.Head.FriendlyName;
                }
            }
            catch { return "unknown"; }
        }
    }
}

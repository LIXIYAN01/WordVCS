using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using WordVCS.Core.Models;
using WordVCS.Core.Services;
using WordVCS.UI;
using WordVCS.UI.Controls;
using WordVCS.UI.ViewModels;

namespace WordVCS.AddIn
{
    /// <summary>
    /// 任务窗格管理器 — 连接 WPF UI 与 Core 服务层的桥梁。
    /// 负责：数据加载、事件响应、Word COM 交互。
    /// 这是一个独立于 VSTO 的纯 WPF 伴侣应用实现。
    /// </summary>
    public class TaskPaneManager
    {
        private readonly WordVCSTaskPane _taskPane;
        private readonly Word.Document _wordDoc;
        private readonly Word.Application _wordApp;
        private readonly IGitService _git;
        private readonly ICommentExtractor _commentExtractor;
        private readonly IDiffService _diff;
        private readonly IMappingService _mapping;
        private readonly IRepositoryManager _repoMgr;
        private readonly WordCommentService _wordCommentSvc;
        private readonly Window _ownerWindow;

        private string _authorName = "Author";
        private string _authorEmail = "author@example.com";

        public TaskPaneManager(
            WordVCSTaskPane taskPane,
            Word.Document wordDoc,
            Word.Application wordApp)
        {
            _taskPane = taskPane;
            _wordDoc = wordDoc;
            _wordApp = wordApp;
            _ownerWindow = Window.GetWindow(taskPane);

            _git = new GitService();
            _commentExtractor = new CommentExtractor();
            _diff = new DiffService();
            _mapping = new MappingService();
            _repoMgr = new RepositoryManager(_git, _commentExtractor, _mapping);
            _wordCommentSvc = new WordCommentService(_commentExtractor);

            // Wire up ViewModel events
            var vm = _taskPane.ViewModel;
            if (vm != null)
            {
                vm.CommitRequested += OnCommitRequested;
                vm.BranchCreateRequested += OnBranchCreateRequested;
                vm.BranchSwitchRequested += OnBranchSwitchRequested;
                vm.TagCreateRequested += OnTagCreateRequested;
                vm.VersionRestoreRequested += OnVersionRestoreRequested;
                vm.VersionDiffRequested += OnVersionDiffRequested;
            }
        }

        public void Initialize()
        {
            LoadConfig();
            Refresh();
        }

        public void Refresh()
        {
            var docPath = GetDocumentPath();
            if (string.IsNullOrEmpty(docPath)) return;

            _repoMgr.EnsureRepository(docPath, _authorName, _authorEmail, "draft")
                .ContinueWith(_ =>
                {
                    Application.Current?.Dispatcher.Invoke(() =>
                    {
                        LoadVersionHistory(docPath);
                        LoadComments(docPath);
                        LoadBranches(docPath);
                        UpdateStatus(docPath);
                    });
                });
        }

        #region Tab Navigation

        public void SwitchToTab(int index)
        {
            _taskPane.SwitchToTab(index);
        }

        public void ShowCommitDialog()
        {
            var dialog = new CommitDialog(_taskPane.ViewModel);
            dialog.Owner = _ownerWindow;
            dialog.ShowDialog();
        }

        public void ShowBranchDialog()
        {
            var dialog = new BranchDialog(_taskPane.ViewModel);
            dialog.Owner = _ownerWindow;
            dialog.ShowDialog();
        }

        public void ShowTagDialog()
        {
            string sha = null;
            if (_taskPane.ViewModel.VersionHistory.Count > 0)
                sha = _taskPane.ViewModel.VersionHistory[0].Sha;

            var dialog = new TagDialog(_taskPane.ViewModel, sha);
            dialog.Owner = _ownerWindow;
            dialog.ShowDialog();
        }

        public void ShowFeedbackImportDialog()
        {
            var dialog = new FeedbackImportDialog(_taskPane.ViewModel);
            dialog.Owner = _ownerWindow;
            if (dialog.ShowDialog() == true && dialog.Tag is string feedbackPath)
            {
                ProcessFeedbackImport(feedbackPath);
            }
        }

        public void ShowSettings()
        {
            MessageBox.Show(
                "设置面板将在后续版本中实现。\n"
                + "您可以在论文目录的 .wordvcs/config.json 中手动修改配置。\n\n"
                + "当前作者: " + _authorName + " <" + _authorEmail + ">",
                "WordVCS 设置",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        #endregion

        #region Event Handlers

        private async void OnCommitRequested(
            string message, string detail, List<string> resolvedCommentIds)
        {
            try
            {
                var docPath = GetDocumentPath();
                if (string.IsNullOrEmpty(docPath)) return;

                // Save document first
                _wordDoc.Save();

                var addressedComments = new List<CommentRecord>();
                foreach (var c in _taskPane.ViewModel.GetCheckedComments())
                {
                    addressedComments.Add(new CommentRecord
                    {
                        Id = c.Id,
                        Author = c.Author,
                        Content = c.Content,
                        CreatedAt = c.CreatedAt,
                        SelectedText = c.SelectedText,
                        Status = CommentStatus.Addressed
                    });
                }

                var sha = await _repoMgr.ExecuteCommit(
                    docPath, message, _authorName, _authorEmail, addressedComments);

                if (!string.IsNullOrEmpty(sha))
                {
                    MessageBox.Show(
                        "提交成功！\n版本: " + sha.Substring(0, 7) + "\n"
                        + "已处理 " + addressedComments.Count + " 条批注",
                        "提交成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    Refresh();
                }
                else
                {
                    MessageBox.Show(
                        "没有检测到文件变更，提交未执行。",
                        "无需提交", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("提交失败: " + ex.Message, "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnBranchCreateRequested(string branchName)
        {
            var docPath = GetDocumentPath();
            _git.CreateBranch(docPath, branchName).ContinueWith(t =>
            {
                Application.Current?.Dispatcher.Invoke(() => Refresh());
            });
        }

        private void OnBranchSwitchRequested(string branchName)
        {
            var docPath = GetDocumentPath();
            _wordDoc.Save();
            _git.SwitchBranch(docPath, branchName).ContinueWith(t =>
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    var docPathReload = GetDocumentPath();
                    _wordDoc.Close();
                    _wordApp.Documents.Open(docPathReload);
                    Refresh();
                });
            });
        }

        private void OnTagCreateRequested(string tagName, string commitSha)
        {
            var docPath = GetDocumentPath();
            _git.CreateTag(docPath, tagName,
                "WordVCS tag: " + tagName).ContinueWith(t =>
                {
                    Application.Current?.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show("标签 '" + tagName + "' 创建成功！",
                            "标签已创建", MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        Refresh();
                    });
                });
        }

        private void OnVersionRestoreRequested(string commitSha)
        {
            var result = MessageBox.Show(
                "确认要恢复版本 " + commitSha?.Substring(0, 7) + " 吗？\n"
                + "当前修改将丢失。建议先保存当前版本。",
                "确认恢复", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var docPath = GetDocumentPath();
                var tempPath = Path.GetTempFileName() + ".docx";

                _git.RestoreVersion(docPath, commitSha, tempPath)
                    .ContinueWith(t =>
                    {
                        if (t.Result)
                        {
                            Application.Current?.Dispatcher.Invoke(() =>
                            {
                                _wordDoc.Close();
                                File.Copy(tempPath, docPath, true);
                                try { File.Delete(tempPath); } catch { }
                                _wordApp.Documents.Open(docPath);
                                Refresh();
                            });
                        }
                    });
            }
        }

        private void OnVersionDiffRequested(string oldSha, string newSha)
        {
            RefreshDiffView();
        }

        #endregion

        #region Data Loading

        private void LoadVersionHistory(string docPath)
        {
            var versions = new List<VersionEntryViewModel>();
            try
            {
                var history = _git.GetHistory(docPath, maxCount: 50);
                foreach (var v in history)
                {
                    var vm = new VersionEntryViewModel
                    {
                        Sha = v.Sha,
                        ShortMessage = v.ShortMessage,
                        FullMessage = v.Message,
                        Author = v.Author,
                        AuthorDateString = v.CommittedAt.ToString("yyyy-MM-dd HH:mm"),
                        CommittedAt = v.CommittedAt,
                    };

                    if (v.Tags != null)
                        foreach (var tag in v.Tags)
                            vm.Tags.Add(tag);

                    var cmts = _mapping.GetCommentsByCommit(docPath, v.Sha);
                    vm.AddressedCommentCount = cmts.Count;
                    versions.Add(vm);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[WordVCS] LoadVersionHistory: " + ex.Message);
            }

            var vm2 = _taskPane.ViewModel;
            vm2.VersionHistory.Clear();
            foreach (var v in versions)
                vm2.VersionHistory.Add(v);
        }

        private void LoadComments(string docPath)
        {
            var comments = new List<CommentEntryViewModel>();
            try
            {
                List<CommentRecord> wordComments;
                try
                {
                    wordComments = _wordCommentSvc.ExtractFromDocument(_wordDoc);
                }
                catch
                {
                    wordComments = _commentExtractor.ExtractFromFile(docPath);
                }

                var persisted = _mapping.GetPersistedComments(docPath);
                var statusMap = persisted.ToDictionary(p => p.Id, p => p.Status);

                foreach (var c in wordComments)
                {
                    var vm = new CommentEntryViewModel
                    {
                        Id = c.Id,
                        Author = c.Author,
                        Content = c.Content,
                        CreatedDateString = c.CreatedAt.ToString("yyyy-MM-dd HH:mm"),
                        SelectedText = c.SelectedText,
                        Status = statusMap.ContainsKey(c.Id)
                            ? statusMap[c.Id].ToString() : "Pending",
                        ResolvedInCommit = c.ResolvedInCommit,
                        CreatedAt = c.CreatedAt
                    };
                    comments.Add(vm);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[WordVCS] LoadComments: " + ex.Message);
            }

            var vm2 = _taskPane.ViewModel;
            vm2.ActiveComments.Clear();
            foreach (var c in comments)
                vm2.ActiveComments.Add(c);
        }

        private void LoadBranches(string docPath)
        {
            var vm = _taskPane.ViewModel;
            vm.Branches.Clear();

            try
            {
                var branches = _git.ListBranches(docPath);
                foreach (var b in branches)
                {
                    vm.Branches.Add(new BranchEntryViewModel
                    {
                        Name = b.Name,
                        IsCurrent = b.IsCurrent,
                        HeadCommitSha = b.HeadCommitSha
                    });
                }
            }
            catch { }
        }

        private void UpdateStatus(string docPath)
        {
            var status = _repoMgr.GetStatus(docPath);
            var vm = _taskPane.ViewModel;
            vm.CurrentBranch = status.CurrentBranch;
            vm.HeadCommit = status.HeadCommitSha;
            vm.IsDirty = status.HasUncommittedChanges;
            vm.UnresolvedCommentCount = status.UnresolvedCommentCount;
            vm.TotalCommentCount = status.TotalCommentCount;
        }

        private void RefreshDiffView()
        {
            try
            {
                var docPath = GetDocumentPath();
                var blocks = new List<DiffBlockViewModel>();
                // Real diff will be computed when user selects versions
                _taskPane.UpdateDiffBlocks(blocks);
            }
            catch { }
        }

        #endregion

        #region Feedback Import

        private async void ProcessFeedbackImport(string feedbackPath)
        {
            try
            {
                var docPath = GetDocumentPath();
                var repoDir = Path.GetDirectoryName(docPath);
                var feedbackName = Path.GetFileName(feedbackPath);
                var destPath = Path.Combine(repoDir, feedbackName);
                File.Copy(feedbackPath, destPath, true);

                int round = 1;
                var tags = _git.ListTags(docPath);
                while (tags.Contains("feedback/r" + round + "-received")) round++;

                var branchName = "feedback/round" + round;
                await _git.CreateBranch(docPath, branchName);
                await _repoMgr.ExecuteCommit(docPath,
                    "[接收反馈] 第" + round + "轮导师批注",
                    _authorName, _authorEmail, null);
                await _git.CreateTag(docPath,
                    "feedback/r" + round + "-received",
                    "收到第" + round + "轮导师反馈");

                var feedbackComments = _commentExtractor.ExtractFromFile(feedbackPath);

                var history = _git.GetHistory(docPath, maxCount: 1);
                var commitSha = history.Count > 0 ? history[0].Sha : "";
                await _mapping.ImportComments(docPath, feedbackComments, commitSha);

                await _git.SwitchBranch(docPath, "draft");

                Application.Current?.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(
                        "第" + round + "轮反馈导入成功！\n"
                        + "已提取 " + feedbackComments.Count + " 条批注\n"
                        + "分支: " + branchName + "\n"
                        + "标签: feedback/r" + round + "-received\n\n"
                        + "请切换到 draft 分支对照批注修改论文。",
                        "导入成功", MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    Refresh();
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("导入失败: " + ex.Message, "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Helpers

        private string GetDocumentPath()
        {
            try { return _wordDoc?.FullName; }
            catch { return null; }
        }

        private void LoadConfig()
        {
            try
            {
                var docPath = GetDocumentPath();
                if (string.IsNullOrEmpty(docPath)) return;

                var configPath = Path.Combine(
                    Path.GetDirectoryName(docPath), ".wordvcs", "config.json");
                if (File.Exists(configPath))
                {
                    var json = File.ReadAllText(configPath);
                    var config = Newtonsoft.Json.JsonConvert
                        .DeserializeObject<RepositoryConfig>(json);
                    if (config != null)
                    {
                        _authorName = config.AuthorName ?? "Author";
                        _authorEmail = config.AuthorEmail ?? "author@example.com";
                    }
                }
            }
            catch { }
        }

        #endregion
    }
}

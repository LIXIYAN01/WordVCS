using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;
using WordVCS.Core.Models;
using WordVCS.Core.Services;
using WordVCS.UI;
using WordVCS.UI.Controls;
using WordVCS.UI.ViewModels;

namespace WordVCS.AddIn
{
    /// <summary>
    /// Bridge between WPF task pane UI and Core services.
    /// Manages data loading, events, and Word COM interaction.
    /// </summary>
    public class TaskPaneManager
    {
        private readonly WordVCSTaskPane _pane;
        private readonly Word.Document _doc;
        private readonly Word.Application _app;
        private readonly IGitService _git = new GitService();
        private readonly ICommentExtractor _ce = new CommentExtractor();
        private readonly IDiffService _diff = new DiffService();
        private readonly IMappingService _map = new MappingService();
        private readonly IRepositoryManager _repo;
        private readonly WordCommentService _wc;

        private string _an = "Author";
        private string _ae = "author@example.com";

        public TaskPaneManager(WordVCSTaskPane pane, Word.Document doc,
            Word.Application app)
        {
            _pane = pane; _doc = doc; _app = app;
            _repo = new RepositoryManager(_git, _ce, _map);
            _wc = new WordCommentService(_ce);
            var vm = _pane.ViewModel;
            vm.CommitRequested += OnCommit;
            vm.BranchCreateRequested += OnBranchCreate;
            vm.BranchSwitchRequested += OnBranchSwitch;
            vm.TagCreateRequested += OnTagCreate;
            vm.VersionRestoreRequested += OnRestore;
        }

        public void Initialize() { LoadCfg(); Refresh(); }

        public void Refresh()
        {
            var dp = DocPath(); if (dp == null) return;
            _repo.EnsureRepository(dp, _an, _ae, "draft").ContinueWith(_ =>
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    LoadHistory(dp); LoadComments(dp);
                    LoadBranches(dp); UpdateStatus(dp);
                });
            });
        }

        public void SwitchToTab(int i) => _pane.SwitchToTab(i);

        public void ShowCommitDialog()
            => new CommitDialog(_pane.ViewModel)
            { WindowStartupLocation = WindowStartupLocation.CenterScreen }
            .ShowDialog();

        public void ShowBranchDialog()
            => new BranchDialog(_pane.ViewModel)
            { WindowStartupLocation = WindowStartupLocation.CenterScreen }
            .ShowDialog();

        public void ShowTagDialog()
        {
            string sha = null;
            if (_pane.ViewModel.VersionHistory.Count > 0)
                sha = _pane.ViewModel.VersionHistory[0].Sha;
            new TagDialog(_pane.ViewModel, sha)
            { WindowStartupLocation = WindowStartupLocation.CenterScreen }
            .ShowDialog();
        }

        public void ShowFeedbackImportDialog()
        {
            var dlg = new FeedbackImportDialog(_pane.ViewModel)
            { WindowStartupLocation = WindowStartupLocation.CenterScreen };
            if (dlg.ShowDialog() == true && dlg.Tag is string fp)
                _ = ImportFeedback(fp);
        }

        public void ShowSettings()
        {
            MessageBox.Show(
                "Settings: edit .wordvcs/config.json in your thesis folder.\n\n" +
                "Author: " + _an + " <" + _ae + ">",
                "WordVCS", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #region Events

        private async void OnCommit(string msg, string detail,
            List<string> resolvedIds)
        {
            try
            {
                var dp = DocPath(); if (dp == null) return;
                _doc.Save();
                var list = new List<CommentRecord>();
                foreach (var c in _pane.ViewModel.GetCheckedComments())
                    list.Add(new CommentRecord
                    {
                        Id = c.Id, Author = c.Author, Content = c.Content,
                        CreatedAt = c.CreatedAt, SelectedText = c.SelectedText,
                        Status = CommentStatus.Addressed
                    });
                var sha = await _repo.ExecuteCommit(dp, msg, _an, _ae, list);
                if (!string.IsNullOrEmpty(sha))
                {
                    MessageBox.Show("Commit OK!\nVersion: " +
                        sha.Substring(0, 7) + "\nResolved: " + list.Count,
                        "WordVCS", MessageBoxButton.OK, MessageBoxImage.Information);
                    Refresh();
                }
                else
                    MessageBox.Show("No changes detected.", "WordVCS",
                        MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Commit failed: " + ex.Message, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnBranchCreate(string name)
        {
            var dp = DocPath();
            _git.CreateBranch(dp, name).ContinueWith(_ =>
                Application.Current?.Dispatcher.Invoke(() => Refresh()));
        }

        private void OnBranchSwitch(string name)
        {
            var dp = DocPath(); _doc.Save();
            _git.SwitchBranch(dp, name).ContinueWith(_ =>
            {
                Application.Current?.Dispatcher.Invoke(() =>
                { _doc.Close(); _app.Documents.Open(dp); Refresh(); });
            });
        }

        private void OnTagCreate(string name, string sha)
        {
            _git.CreateTag(DocPath(), name, "Tag: " + name).ContinueWith(_ =>
            {
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Tag '" + name + "' created.", "WordVCS",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    Refresh();
                });
            });
        }

        private void OnRestore(string sha)
        {
            if (MessageBox.Show("Restore version " +
                sha?.Substring(0, 7) + "?", "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Warning)
                != MessageBoxResult.Yes) return;
            var dp = DocPath();
            var tp = Path.GetTempFileName() + ".docx";
            _git.RestoreVersion(dp, sha, tp).ContinueWith(t =>
            {
                if (!t.Result) return;
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    _doc.Close(); File.Copy(tp, dp, true);
                    try { File.Delete(tp); } catch { }
                    _app.Documents.Open(dp); Refresh();
                });
            });
        }

        #endregion

        #region Data

        private void LoadHistory(string dp)
        {
            var l = new List<VersionEntryViewModel>();
            try
            {
                foreach (var v in _git.GetHistory(dp, maxCount: 50))
                {
                    var vm = new VersionEntryViewModel
                    {
                        Sha = v.Sha, ShortMessage = v.ShortMessage,
                        FullMessage = v.Message, Author = v.Author,
                        AuthorDateString = v.CommittedAt.ToString("yyyy-MM-dd HH:mm"),
                        CommittedAt = v.CommittedAt,
                        AddressedCommentCount =
                            _map.GetCommentsByCommit(dp, v.Sha).Count
                    };
                    if (v.Tags != null)
                        foreach (var t in v.Tags) vm.Tags.Add(t);
                    l.Add(vm);
                }
            }
            catch { }
            var pv = _pane.ViewModel;
            pv.VersionHistory.Clear();
            foreach (var v in l) pv.VersionHistory.Add(v);
        }

        private void LoadComments(string dp)
        {
            var l = new List<CommentEntryViewModel>();
            try
            {
                List<CommentRecord> wc;
                try { wc = _wc.ExtractFromDocument(_doc); }
                catch { wc = _ce.ExtractFromFile(dp); }
                var sm = _map.GetPersistedComments(dp)
                    .ToDictionary(p => p.Id, p => p.Status);
                foreach (var c in wc)
                    l.Add(new CommentEntryViewModel
                    {
                        Id = c.Id, Author = c.Author, Content = c.Content,
                        CreatedDateString = c.CreatedAt.ToString("yyyy-MM-dd HH:mm"),
                        SelectedText = c.SelectedText,
                        Status = sm.ContainsKey(c.Id) ? sm[c.Id].ToString() : "Pending",
                        CreatedAt = c.CreatedAt
                    });
            }
            catch { }
            var pv = _pane.ViewModel;
            pv.ActiveComments.Clear();
            foreach (var c in l) pv.ActiveComments.Add(c);
        }

        private void LoadBranches(string dp)
        {
            var vm = _pane.ViewModel;
            vm.Branches.Clear();
            try
            {
                foreach (var b in _git.ListBranches(dp))
                    vm.Branches.Add(new BranchEntryViewModel
                    {
                        Name = b.Name, IsCurrent = b.IsCurrent,
                        HeadCommitSha = b.HeadCommitSha
                    });
            }
            catch { }
        }

        private void UpdateStatus(string dp)
        {
            var s = _repo.GetStatus(dp);
            var vm = _pane.ViewModel;
            vm.CurrentBranch = s.CurrentBranch ?? "?";
            vm.HeadCommit = s.HeadCommitSha ?? "";
            vm.IsDirty = s.HasUncommittedChanges;
            vm.UnresolvedCommentCount = s.UnresolvedCommentCount;
            vm.TotalCommentCount = s.TotalCommentCount;
        }

        #endregion

        #region Feedback

        private async System.Threading.Tasks.Task ImportFeedback(string fp)
        {
            try
            {
                var dp = DocPath();
                var dir = Path.GetDirectoryName(dp);
                var dest = Path.Combine(dir, Path.GetFileName(fp));
                File.Copy(fp, dest, true);
                int r = 1;
                var tags = _git.ListTags(dp);
                while (tags.Contains("feedback/r" + r + "-received")) r++;
                var bn = "feedback/round" + r;
                await _git.CreateBranch(dp, bn);
                await _repo.ExecuteCommit(dp,
                    "[Import] Advisor feedback round " + r, _an, _ae, null);
                await _git.CreateTag(dp, "feedback/r" + r + "-received",
                    "Round " + r + " feedback");
                var fbc = _ce.ExtractFromFile(fp);
                var h = _git.GetHistory(dp, maxCount: 1);
                await _map.ImportComments(dp, fbc,
                    h.Count > 0 ? h[0].Sha : "");
                await _git.SwitchBranch(dp, "draft");
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Round " + r + " imported!\n" +
                        fbc.Count + " comments.\nBranch: " + bn,
                        "WordVCS", MessageBoxButton.OK, MessageBoxImage.Information);
                    Refresh();
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed: " + ex.Message, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Helpers

        private string DocPath()
        { try { return _doc?.FullName; } catch { return null; } }

        private void LoadCfg()
        {
            try
            {
                var dp = DocPath(); if (dp == null) return;
                var cp = Path.Combine(Path.GetDirectoryName(dp),
                    ".wordvcs", "config.json");
                if (File.Exists(cp))
                {
                    var cfg = JsonConvert.DeserializeObject<RepositoryConfig>(
                        File.ReadAllText(cp));
                    if (cfg != null)
                    { _an = cfg.AuthorName ?? "Author";
                      _ae = cfg.AuthorEmail ?? "author@example.com"; }
                }
            }
            catch { }
        }

        #endregion
    }
}

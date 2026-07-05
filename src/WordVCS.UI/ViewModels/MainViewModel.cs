using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace WordVCS.UI.ViewModels
{
    /// <summary>
    /// 主视图模型 — 侧边栏的所有数据绑定源
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _currentBranch = "draft";
        private string _headCommit = "";
        private int _unresolvedCount;
        private int _totalCount;
        private bool _isDirty;
        private int _selectedTabIndex;

        public string CurrentBranch
        {
            get => _currentBranch;
            set { _currentBranch = value; OnPropertyChanged(); }
        }

        public string HeadCommit
        {
            get => _headCommit;
            set { _headCommit = value; OnPropertyChanged(); }
        }

        public int UnresolvedCommentCount
        {
            get => _unresolvedCount;
            set { _unresolvedCount = value; OnPropertyChanged(); }
        }

        public int TotalCommentCount
        {
            get => _totalCount;
            set { _totalCount = value; OnPropertyChanged(); }
        }

        public bool IsDirty
        {
            get => _isDirty;
            set { _isDirty = value; OnPropertyChanged(); }
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set { _selectedTabIndex = value; OnPropertyChanged(); }
        }

        public string StatusText =>
            IsDirty ? "有未提交的修改" : "工作区干净";

        // Bindable collections
        public ObservableCollection<VersionEntryViewModel> VersionHistory { get; }
            = new ObservableCollection<VersionEntryViewModel>();

        public ObservableCollection<CommentEntryViewModel> ActiveComments { get; }
            = new ObservableCollection<CommentEntryViewModel>();

        public ObservableCollection<DiffBlockViewModel> DiffBlocks { get; }
            = new ObservableCollection<DiffBlockViewModel>();

        public ObservableCollection<BranchEntryViewModel> Branches { get; }
            = new ObservableCollection<BranchEntryViewModel>();

        // Events for AddIn layer to hook into
        public event Action<string, string, List<string>> CommitRequested;
        public event Action<string> BranchSwitchRequested;
        public event Action<string> BranchCreateRequested;
        public event Action<string, string> TagCreateRequested;
        public event Action FeedbackImportRequested;
        public event Action SettingsRequested;
        public event Action<string> VersionRestoreRequested;
        public event Action<string, string> VersionDiffRequested;

        public void RequestCommit(string summary, string detail,
            List<string> resolvedCommentIds)
        {
            CommitRequested?.Invoke(summary, detail, resolvedCommentIds);
        }

        public void RequestBranchSwitch(string branchName)
            => BranchSwitchRequested?.Invoke(branchName);

        public void RequestBranchCreate(string branchName)
            => BranchCreateRequested?.Invoke(branchName);

        public void RequestTagCreate(string tagName, string commitSha)
            => TagCreateRequested?.Invoke(tagName, commitSha);

        public void RequestFeedbackImport()
            => FeedbackImportRequested?.Invoke();

        public void RequestSettings()
            => SettingsRequested?.Invoke();

        public void RequestVersionRestore(string commitSha)
            => VersionRestoreRequested?.Invoke(commitSha);

        public void RequestVersionDiff(string oldSha, string newSha)
            => VersionDiffRequested?.Invoke(oldSha, newSha);

        /// <summary>
        /// 获取当前所有被勾选的批注
        /// </summary>
        public List<CommentEntryViewModel> GetCheckedComments()
        {
            var result = new List<CommentEntryViewModel>();
            foreach (var c in ActiveComments)
                if (c.IsChecked) result.Add(c);
            return result;
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 版本历史条目 ViewModel
    /// </summary>
    public class VersionEntryViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public string Sha { get; set; }
        public string ShortSha => !string.IsNullOrEmpty(Sha) && Sha.Length >= 7
            ? Sha.Substring(0, 7) : Sha;
        public string ShortMessage { get; set; }
        public string FullMessage { get; set; }
        public string Author { get; set; }
        public string AuthorDateString { get; set; }
        public DateTime CommittedAt { get; set; }
        public ObservableCollection<string> Tags { get; set; }
            = new ObservableCollection<string>();
        public int AddressedCommentCount { get; set; }
        public bool IsHead { get; set; }

        public string DisplayTags => Tags.Count > 0
            ? string.Join(", ", Tags) : "";
        public bool HasTags => Tags.Count > 0;
        public bool HasAddressedComments => AddressedCommentCount > 0;

        public string CommitSummary =>
            (IsHead ? "● " : "○ ") + ShortMessage;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 批注条目 ViewModel
    /// </summary>
    public class CommentEntryViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isChecked;

        public string Id { get; set; }
        public string Author { get; set; }
        public string Content { get; set; }
        public string CreatedDateString { get; set; }
        public string SelectedText { get; set; }
        public string Status { get; set; } = "Pending";
        public string ResolvedInCommit { get; set; }
        public DateTime CreatedAt { get; set; }

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(); }
        }

        public bool IsPending => Status == "Pending";
        public bool IsAddressed => Status == "Addressed";
        public bool IsDismissed => Status == "Dismissed";

        public string StatusDisplay
        {
            get
            {
                switch (Status)
                {
                    case "Addressed": return "✔ 已处理";
                    case "Dismissed": return "✖ 已忽略";
                    case "Replied": return "↩ 已回复";
                    default: return "";
                }
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 差异块 ViewModel
    /// </summary>
    public class DiffBlockViewModel
    {
        public string Text { get; set; }
        public string Indicator { get; set; }
        public string BgColor { get; set; }
        public string TextColor { get; set; }
        public string Type { get; set; } // Added/Removed/Modified/Unchanged
    }

    /// <summary>
    /// 分支条目 ViewModel
    /// </summary>
    public class BranchEntryViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public string Name { get; set; }
        public bool IsCurrent { get; set; }
        public string HeadCommitSha { get; set; }

        public string DisplayName => IsCurrent ? $"● {Name} (当前)" : $"  {Name}";

        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}

using System.Windows;
using System.Windows.Controls;
using WordVCS.UI.Controls;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI
{
    /// <summary>
    /// WordVCSTaskPane.xaml 的交互逻辑。
    /// 这是嵌入 Word 自定义任务窗格的 WPF 用户控件。
    /// </summary>
    public partial class WordVCSTaskPane : UserControl
    {
        public MainViewModel ViewModel => DataContext as MainViewModel;

        public WordVCSTaskPane()
        {
            InitializeComponent();
            DataContext = new MainViewModel();

            // Wire up events
            if (ViewModel != null)
            {
                ViewModel.CommitRequested += OnViewModelCommitRequested;
                ViewModel.BranchSwitchRequested += OnViewModelBranchSwitchRequested;
                ViewModel.FeedbackImportRequested += OnViewModelFeedbackImportRequested;
                ViewModel.SettingsRequested += OnViewModelSettingsRequested;
                ViewModel.TagCreateRequested += OnViewModelTagCreateRequested;
                ViewModel.VersionRestoreRequested += OnViewModelVersionRestoreRequested;
            }
        }

        // ---- External hooks for AddIn layer ----

        /// <summary>
        /// 将服务层数据注入 ViewModel。
        /// 由 AddIn 层的 TaskPaneManager 在初始化时调用。
        /// </summary>
        public void LoadData(
            System.Collections.Generic.List<VersionEntryViewModel> versions,
            System.Collections.Generic.List<CommentEntryViewModel> comments,
            string currentBranch,
            string headCommit,
            bool hasUncommittedChanges)
        {
            if (ViewModel == null) return;

            ViewModel.VersionHistory.Clear();
            foreach (var v in versions)
                ViewModel.VersionHistory.Add(v);

            ViewModel.ActiveComments.Clear();
            foreach (var c in comments)
                ViewModel.ActiveComments.Add(c);

            ViewModel.CurrentBranch = currentBranch;
            ViewModel.HeadCommit = headCommit;
            ViewModel.IsDirty = hasUncommittedChanges;
        }

        public void UpdateDiffBlocks(
            System.Collections.Generic.List<DiffBlockViewModel> blocks)
        {
            if (ViewModel == null) return;
            ViewModel.DiffBlocks.Clear();
            foreach (var b in blocks)
                ViewModel.DiffBlocks.Add(b);
        }

        public void SwitchToTab(int index)
        {
            if (ViewModel != null)
                ViewModel.SelectedTabIndex = index;
        }

        // ---- Button handlers ----

        private void OnCommitClick(object sender, RoutedEventArgs e)
        {
            var dialog = new CommitDialog(ViewModel);
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();
        }

        private void OnBranchClick(object sender, RoutedEventArgs e)
        {
            var dialog = new BranchDialog(ViewModel);
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();
        }

        private void OnTagClick(object sender, RoutedEventArgs e)
        {
            // Determine selected commit from version history
            string selectedSha = null;
            if (ViewModel.VersionHistory.Count > 0)
                selectedSha = ViewModel.VersionHistory[0].Sha;

            var dialog = new TagDialog(ViewModel, selectedSha);
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();
        }

        private void OnImportFeedbackClick(object sender, RoutedEventArgs e)
        {
            ViewModel?.RequestFeedbackImport();
        }

        private void OnSettingsClick(object sender, RoutedEventArgs e)
        {
            ViewModel?.RequestSettings();
        }

        // ---- ViewModel event relay ----

        private void OnViewModelCommitRequested(
            string summary, string detail,
            System.Collections.Generic.List<string> resolvedIds)
        {
            // Handled internally by CommitDialog
        }

        private void OnViewModelBranchSwitchRequested(string branchName)
        {
            // Handled by BranchDialog
        }

        private void OnViewModelFeedbackImportRequested()
        {
            var dialog = new FeedbackImportDialog(ViewModel);
            dialog.Owner = Window.GetWindow(this);
            dialog.ShowDialog();
        }

        private void OnViewModelSettingsRequested()
        {
            // Will show settings dialog
            MessageBox.Show("设置面板将在后续版本中提供。\n"
                + "您可以在 .wordvcs/config.json 中手动修改配置。",
                "WordVCS 设置", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void OnViewModelTagCreateRequested(
            string tagName, string commitSha)
        {
            // Handled by TagDialog
        }

        private void OnViewModelVersionRestoreRequested(string commitSha)
        {
            var result = MessageBox.Show(
                $"确认要恢复版本 {commitSha?.Substring(0, 7)} 吗？\n"
                + "当前未保存的修改将丢失。",
                "确认恢复", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            // Actual restore handled by AddIn layer via event
        }
    }
}

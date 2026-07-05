using System.Collections.Generic;
using System.Linq;
using System.Windows;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI.Controls
{
    public partial class CommitDialog : Window
    {
        private readonly MainViewModel _viewModel;

        public CommitDialog(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = new CommitDialogViewModel(viewModel);
            LoadAddressedComments();
        }

        private void LoadAddressedComments()
        {
            var checkedComments = _viewModel.GetCheckedComments();
            AddressedCommentsList.ItemsSource = checkedComments;
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnSubmitClick(object sender, RoutedEventArgs e)
        {
            var summary = SummaryBox.Text?.Trim();
            if (string.IsNullOrEmpty(summary))
            {
                MessageBox.Show("请输入提交摘要。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var commitType =
                ((System.Windows.Controls.ComboBoxItem)CommitTypeCombo.SelectedItem)
                ?.Content?.ToString()?.Split(']')[0] + "]" ?? "[修改]";
            var detail = DetailBox.Text?.Trim() ?? "";
            var fullMessage = $"{commitType} {summary}";
            if (!string.IsNullOrEmpty(detail))
                fullMessage += "\n\n" + detail;

            var checkedComments = _viewModel.GetCheckedComments();
            var resolvedIds = checkedComments
                .Select(c => c.Id).ToList();

            // Add addressed comments info to message
            if (resolvedIds.Count > 0)
            {
                fullMessage += "\n\n关联批注:";
                foreach (var c in checkedComments)
                {
                    var contentLen = System.Math.Min(c.Content?.Length ?? 0, 80);
                    var snippet = c.Content?.Substring(0, contentLen) ?? "";
                    fullMessage += $"\n  - {c.Id}: {snippet}";
                }
            }

            _viewModel.RequestCommit(fullMessage, detail, resolvedIds);

            DialogResult = true;
            Close();
        }
    }

    /// <summary>
    /// Simple ViewModel for the commit dialog
    /// </summary>
    internal class CommitDialogViewModel
    {
        private readonly MainViewModel _parent;

        public CommitDialogViewModel(MainViewModel parent)
        {
            _parent = parent;
        }

        public string CommentCountText
        {
            get
            {
                var count = _parent.GetCheckedComments().Count;
                return count > 0
                    ? $"{count} 条批注，将在提交后标记为已处理"
                    : "无勾选的批注";
            }
        }
    }
}

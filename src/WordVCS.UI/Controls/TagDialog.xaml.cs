using System.Windows;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI.Controls
{
    public partial class TagDialog : Window
    {
        private readonly MainViewModel _viewModel;
        private readonly string _commitSha;

        public TagDialog(MainViewModel viewModel, string commitSha)
        {
            InitializeComponent();
            _viewModel = viewModel;
            _commitSha = commitSha;
            CommitShaText.Text = commitSha?.Length > 7
                ? commitSha.Substring(0, 7) : commitSha ?? "(HEAD)";
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnCreate(object sender, RoutedEventArgs e)
        {
            var tagName = TagNameBox.Text?.Trim();
            if (string.IsNullOrEmpty(tagName))
            {
                MessageBox.Show("请输入标签名称。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var message = TagMessageBox.Text?.Trim() ?? "";
            _viewModel.RequestTagCreate(tagName, _commitSha);
            DialogResult = true;
            Close();
        }
    }
}

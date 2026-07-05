using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI.Controls
{
    public partial class BranchDialog : Window
    {
        private readonly MainViewModel _viewModel;

        public BranchDialog(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
        }

        private void OnCreateBranch(object sender, RoutedEventArgs e)
        {
            var name = NewBranchName.Text?.Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("请输入分支名称。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            _viewModel.RequestBranchCreate(name);
            DialogResult = true;
            Close();
        }

        private void OnBranchClicked(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.DataContext is BranchEntryViewModel branch)
            {
                var result = MessageBox.Show(
                    $"切换到分支 '{branch.Name}'？\n当前未保存的修改将被暂存。",
                    "切换分支", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    _viewModel.RequestBranchSwitch(branch.Name);
                    DialogResult = true;
                    Close();
                }
            }
        }

        private void OnClose(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

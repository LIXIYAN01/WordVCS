using System.Windows;
using Microsoft.Win32;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI.Controls
{
    public partial class FeedbackImportDialog : Window
    {
        private readonly MainViewModel _viewModel;

        public FeedbackImportDialog(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
        }

        private void OnBrowse(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*",
                Title = "选择导师反馈文件"
            };

            if (dialog.ShowDialog() == true)
            {
                FilePathBox.Text = dialog.FileName;
            }
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnImport(object sender, RoutedEventArgs e)
        {
            var filePath = FilePathBox.Text?.Trim();
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("请选择导师反馈文件。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("文件不存在，请检查路径。", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Store the path for the AddIn layer to process
            Tag = filePath;
            _viewModel.RequestFeedbackImport();
            DialogResult = true;
            Close();
        }
    }
}

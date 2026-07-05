using System.Windows;
using System.Windows.Controls;

namespace WordVCS.UI.Controls
{
    public partial class DiffViewer : UserControl
    {
        public DiffViewer()
        {
            InitializeComponent();
        }

        private void OnCompareTargetChanged(object sender,
            SelectionChangedEventArgs e)
        {
            // Triggered by AddIn layer to refresh the diff
        }

        private void OnRefreshClick(object sender, RoutedEventArgs e)
        {
            var vm = DataContext as ViewModels.MainViewModel;
            vm?.RequestVersionDiff("HEAD^", "WORKDIR");
        }
    }
}

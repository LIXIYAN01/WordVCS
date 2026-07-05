using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WordVCS.UI.ViewModels;

namespace WordVCS.UI.Controls
{
    public partial class CommentPanel : UserControl
    {
        public CommentPanel()
        {
            InitializeComponent();
        }

        private MainViewModel ViewModel => DataContext as MainViewModel;

        private void OnSelectAll(object sender, RoutedEventArgs e)
        {
            if (ViewModel == null) return;
            foreach (var c in ViewModel.ActiveComments)
                c.IsChecked = true;
        }

        private void OnDeselectAll(object sender, RoutedEventArgs e)
        {
            if (ViewModel == null) return;
            foreach (var c in ViewModel.ActiveComments)
                c.IsChecked = false;
        }

        private void OnFilterChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ViewModel == null) return;
            // Filter is applied via the collection directly by reloading
            // from the AddIn layer. For now, the ComboBox is informational.
        }
    }
}

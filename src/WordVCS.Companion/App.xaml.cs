using System;
using System.Windows;
using System.Windows.Threading;

namespace WordVCS.Companion
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            DispatcherUnhandledException += (s, args) =>
            {
                MessageBox.Show("Error: " + args.Exception.Message + "\n\n"
                    + args.Exception.InnerException?.Message,
                    "WordVCS", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };
        }
    }
}

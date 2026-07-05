using System;
using System.Windows;
using System.Windows.Threading;

namespace WordVCS.AddIn
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Global exception handlers to prevent silent crashes
            DispatcherUnhandledException += (s, args) =>
            {
                MessageBox.Show(
                    "WordVCS 运行时错误:\n\n" + args.Exception.Message
                    + "\n\n" + (args.Exception.InnerException?.Message ?? ""),
                    "WordVCS 错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };

            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                var ex = args.ExceptionObject as Exception;
                MessageBox.Show(
                    "WordVCS 严重错误:\n\n" + (ex?.Message ?? "未知错误"),
                    "WordVCS 错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            };
        }
    }
}

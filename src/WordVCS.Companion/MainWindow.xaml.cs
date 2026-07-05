using System;
using System.Runtime.InteropServices;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace WordVCS.Companion
{
    /// <summary>
    /// WordVCS Companion — standalone app that auto-detects
    /// Word or WPS and connects via COM automation.
    /// Works with both Microsoft Word and WPS Office Writer.
    /// </summary>
    public partial class MainWindow : Window
    {
        private Word.Application _wordApp;
        private Word.Document _activeDoc;
        private UI.WordVCSTaskPane _taskPane;
        private AddIn.TaskPaneManager _manager;
        private string _appType = "";

        public MainWindow()
        {
            try { InitializeComponent(); }
            catch (Exception ex)
            {
                MessageBox.Show("UI init failed: " + ex.Message, "WordVCS",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _taskPane = new UI.WordVCSTaskPane();
            TaskPaneHost.Content = _taskPane;
            Connect();
        }

        private void Connect()
        {
            // Try Word first
            try
            {
                _wordApp = (Word.Application)
                    Marshal.GetActiveObject("Word.Application");
                _appType = "MS Word";
            }
            catch
            {
                // Try WPS Writer
                try
                {
                    _wordApp = (Word.Application)
                        Marshal.GetActiveObject("WPS.Application");
                    _appType = "WPS Writer";
                }
                catch
                {
                    // Try Kingsoft WPS
                    try
                    {
                        _wordApp = (Word.Application)
                            Marshal.GetActiveObject("KWPS.Application");
                        _appType = "WPS Writer";
                    }
                    catch { }
                }
            }

            if (_wordApp != null)
            {
                _wordApp.DocumentChange += OnDocChange;
                _activeDoc = _wordApp.ActiveDocument;
                if (_activeDoc != null) InitDoc(_activeDoc);
                else TxtStatus.Text = _appType + " connected - open a .docx file";
            }
            else
            {
                TxtStatus.Text = "No Word or WPS detected. " +
                    "Please open Word/WPS first.";
            }
        }

        private void OnDocChange()
        {
            Dispatcher.Invoke(() =>
            {
                try
                {
                    _activeDoc = _wordApp?.ActiveDocument;
                    if (_activeDoc != null) InitDoc(_activeDoc);
                }
                catch { }
            });
        }

        private void InitDoc(Word.Document doc)
        {
            try
            {
                var name = System.IO.Path.GetFileName(doc.FullName);
                TxtStatus.Text = "[" + _appType + "] " + name;
            }
            catch
            {
                TxtStatus.Text = "Please save document first (Ctrl+S)";
                return;
            }

            try
            {
                _manager = new AddIn.TaskPaneManager(
                    _taskPane, doc, _wordApp);
                _manager.Initialize();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Init error: " + ex.Message;
            }
        }

        private void OnRefresh(object sender, RoutedEventArgs e)
        {
            if (_wordApp == null) Connect();
            else
            {
                try
                {
                    _activeDoc = _wordApp.ActiveDocument;
                    if (_activeDoc != null) InitDoc(_activeDoc);
                    _manager?.Refresh();
                }
                catch { Connect(); }
            }
        }

        private void OnImport(object sender, RoutedEventArgs e)
        {
            _manager?.ShowFeedbackImportDialog();
        }

        private void OnClosing(object sender,
            System.ComponentModel.CancelEventArgs e)
        {
            if (_wordApp != null)
            {
                try { _wordApp.DocumentChange -= OnDocChange; }
                catch { }
                Marshal.ReleaseComObject(_wordApp);
            }
        }
    }
}

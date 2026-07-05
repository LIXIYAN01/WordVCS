using System;
using System.Runtime.InteropServices;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using WordVCS.UI;

namespace WordVCS.AddIn
{
    public partial class MainWindow : Window
    {
        private Word.Application _wordApp;
        private Word.Document _activeDoc;
        private WordVCSTaskPane _taskPane;
        private TaskPaneManager _manager;

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                _taskPane = new WordVCSTaskPane();
                TaskPaneHost.Content = _taskPane;
            }
            catch (Exception ex)
            {
                TxtDocStatus.Text = "UI 加载失败: " + ex.Message;
                return;
            }

            ConnectToWord();
        }

        private void ConnectToWord()
        {
            try
            {
                _wordApp = (Word.Application)
                    Marshal.GetActiveObject("Word.Application");
                _wordApp.DocumentChange += OnWordDocumentChange;

                _activeDoc = _wordApp.ActiveDocument;
                if (_activeDoc != null)
                    InitializeForDocument(_activeDoc);
                else
                {
                    TxtTitle.Text = "WordVCS (等待文档)";
                    TxtDocStatus.Text = "Word 已连接 — 请打开一个 .docx 论文文件";
                }
            }
            catch (COMException)
            {
                TxtTitle.Text = "WordVCS (等待 Word)";
                TxtDocStatus.Text =
                    "未检测到运行中的 Word\n请先打开 Word 和论文，然后点击 ↻ 刷新";
            }
            catch (Exception ex)
            {
                TxtDocStatus.Text = "连接 Word 失败: " + ex.Message;
            }
        }

        private void OnWordDocumentChange()
        {
            Dispatcher.Invoke(() =>
            {
                try
                {
                    _activeDoc = _wordApp?.ActiveDocument;
                    if (_activeDoc != null)
                        InitializeForDocument(_activeDoc);
                }
                catch { }
            });
        }

        private void InitializeForDocument(Word.Document doc)
        {
            if (doc == null) return;
            try
            {
                var docName = System.IO.Path.GetFileName(doc.FullName);
                TxtTitle.Text = "WordVCS";
                TxtDocStatus.Text = "\U0001f4c4 " + docName;
            }
            catch
            {
                TxtDocStatus.Text = "请先保存文档 (Ctrl+S)";
                return;
            }

            try
            {
                _manager = new TaskPaneManager(_taskPane, doc, _wordApp);
                _manager.Initialize();
            }
            catch (Exception ex)
            {
                TxtDocStatus.Text = "初始化失败: " + ex.Message;
            }
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            // Window is now visible; nothing extra needed
        }

        private void OnRefreshClick(object sender, RoutedEventArgs e)
        {
            if (_wordApp == null)
            {
                ConnectToWord();
                return;
            }
            try
            {
                _activeDoc = _wordApp.ActiveDocument;
                if (_activeDoc != null)
                {
                    InitializeForDocument(_activeDoc);
                    _manager?.Refresh();
                }
            }
            catch { ConnectToWord(); }
        }

        private void OnImportFeedbackClick(object sender, RoutedEventArgs e)
        {
            if (_manager == null)
            {
                MessageBox.Show("请先打开 Word 中的论文文档，再点击刷新。",
                    "WordVCS", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            _manager.ShowFeedbackImportDialog();
        }

        private void OnSettingsClick(object sender, RoutedEventArgs e)
        {
            _manager?.ShowSettings();
        }

        private void OnWindowClosing(object sender,
            System.ComponentModel.CancelEventArgs e)
        {
            if (_wordApp != null)
            {
                try { _wordApp.DocumentChange -= OnWordDocumentChange; }
                catch { }
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }
        }
    }
}

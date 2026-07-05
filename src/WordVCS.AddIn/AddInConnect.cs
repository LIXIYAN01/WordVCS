using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordVCS.AddIn
{
    /// <summary>
    /// Word COM Add-in 连接器 — 实现 IDTExtensibility2 + IRibbonExtensibility
    /// Word 启动时自动加载，在「开始」旁显示「论文版本」标签页
    /// </summary>
    [ComVisible(true)]
    [ProgId("WordVCS.Connect")]
    [Guid("B4E1C2D3-A5F6-7890-ABCD-EF1234567890")]
    public class AddInConnect : IExtensibility, Office.IRibbonExtensibility
    {
        private Word.Application _wordApp;
        private UI.WordVCSTaskPane _wpfControl;
        private TaskPaneManager _manager;
        private UserControl _hostControl;

        #region IExtensibility (IDTExtensibility2)

        public void OnConnection(object application,
            ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _wordApp = application as Word.Application;
                if (_wordApp == null) return;

                // Ensure WPF Application exists on this STA thread
                if (System.Windows.Application.Current == null)
                    new System.Windows.Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

                _wordApp.DocumentOpen += OnDocumentOpen;

                if (_wordApp.Documents.Count > 0 && _wordApp.ActiveDocument != null)
                    InitializeForDocument(_wordApp.ActiveDocument);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordVCS] OnConnection: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_wordApp != null)
                    _wordApp.DocumentOpen -= OnDocumentOpen;
                RemoveTaskPane();
                _wordApp = null;
            }
            catch { }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return RibbonXml.Value;
        }

        #endregion

        #region Document

        private void OnDocumentOpen(Word.Document doc)
        {
            InitializeForDocument(doc);
        }

        public void InitializeForDocument(Word.Document doc)
        {
            if (doc == null) return;
            try
            {
                var path = GetDocPath(doc);
                if (string.IsNullOrEmpty(path)) return;

                RemoveTaskPane();

                _wpfControl = new UI.WordVCSTaskPane();
                _hostControl = new UserControl();
                var elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _wpfControl
                };
                _hostControl.Controls.Add(elementHost);
                _hostControl.CreateControl();
                _hostControl.Width = 360;
                _hostControl.Height = 600;

                _manager = new TaskPaneManager(_wpfControl, doc, _wordApp);
                _manager.Initialize();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordVCS] InitForDoc: " + ex.Message);
            }
        }

        public void ShowCommitDialog() => _manager?.ShowCommitDialog();
        public void ShowBranchDialog() => _manager?.ShowBranchDialog();
        public void ShowTagDialog() => _manager?.ShowTagDialog();
        public void ShowFeedbackDialog() => _manager?.ShowFeedbackImportDialog();
        public void SwitchToTab(int i) => _manager?.SwitchToTab(i);
        public void Refresh() => _manager?.Refresh();
        public Control GetTaskPaneControl() => _hostControl;

        private void RemoveTaskPane()
        {
            _hostControl?.Dispose();
            _hostControl = null;
            _wpfControl = null;
            _manager = null;
        }

        private static string GetDocPath(Word.Document doc)
        {
            try { return doc.FullName; }
            catch { return null; }
        }

        #endregion

        #region Ribbon Callbacks

        public void OnCommit(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            ShowCommitDialog();
        }

        public void OnHistory(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            SwitchToTab(0);
        }

        public void OnDiff(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            SwitchToTab(2);
        }

        public void OnBranch(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            ShowBranchDialog();
        }

        public void OnTag(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            ShowTagDialog();
        }

        public void OnFeedback(Office.IRibbonControl control)
        {
            if (_wordApp?.ActiveDocument != null)
                InitializeForDocument(_wordApp.ActiveDocument);
            ShowFeedbackDialog();
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            _manager?.ShowSettings();
        }

        public void OnShowPanel(Office.IRibbonControl control)
        {
            // Toggle the WPF control panel window
            if (_panelWindow != null && _panelWindow.IsVisible)
            {
                _panelWindow.Hide();
            }
            else
            {
                ShowControlPanel();
            }
        }

        #endregion

        #region Control Panel (WPF Tool Window)

        private Window _panelWindow;

        private void ShowControlPanel()
        {
            if (_wpfControl == null)
            {
                if (_wordApp?.ActiveDocument == null) return;
                InitializeForDocument(_wordApp.ActiveDocument);
            }

            if (_panelWindow == null)
            {
                _panelWindow = new Window
                {
                    Title = "WordVCS - 论文版本控制",
                    Width = 380,
                    Height = 700,
                    MinWidth = 340,
                    MinHeight = 500,
                    WindowStartupLocation = WindowStartupLocation.Manual,
                    Topmost = false,
                    ShowInTaskbar = true,
                    Background = System.Windows.Media.Brushes.White
                };

                // Position to right of Word window
                try
                {
                    var wordHwnd = new System.Windows.Interop.WindowInteropHelper(
                        System.Windows.Application.Current?.MainWindow ?? _panelWindow);
                    // Default position: right side of screen
                    _panelWindow.Left = SystemParameters.WorkArea.Width - 400;
                    _panelWindow.Top = 50;
                }
                catch { }

                _panelWindow.Closing += (s, e) =>
                {
                    e.Cancel = true;
                    _panelWindow.Hide();
                };
            }

            _panelWindow.Content = _hostControl;
            _panelWindow.Show();
            _panelWindow.Activate();
        }

        #endregion
    }

    #region COM Interface Definitions (avoid external Extensibility.dll dependency)

    /// <summary>IDTExtensibility2 COM interface</summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8A-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExtensibility
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    /// <summary>Connection mode enum</summary>
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    /// <summary>Disconnection mode enum</summary>
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    #endregion

    /// <summary>
    /// Embedded Ribbon XML resource
    /// </summary>
    internal static class RibbonXml
    {
        public static readonly string Value = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabWordVCS"" label=""论文版本"" insertBeforeMso=""TabHome"">
        <group id=""grpCommit"" label=""版本控制"">
          <button id=""btnShowPanel"" label=""显示面板"" size=""large""
                  onAction=""OnShowPanel"" imageMso=""ReviewShowAllComments""/>
          <separator id=""s0""/>
          <button id=""btnCommit"" label=""提交新版本"" size=""normal""
                  onAction=""OnCommit"" imageMso=""FileSave""/>
          <button id=""btnHistory"" label=""版本历史"" size=""normal""
                  onAction=""OnHistory"" imageMso=""HistoryStore""/>
          <button id=""btnDiff"" label=""版本对比"" size=""normal""
                  onAction=""OnDiff"" imageMso=""CompareDocuments""/>
        </group>
        <group id=""grpBranch"" label=""分支与标签"">
          <button id=""btnBranch"" label=""分支管理"" size=""large""
                  onAction=""OnBranch"" imageMso=""MergeCells""/>
          <button id=""btnTag"" label=""添加标签"" size=""normal""
                  onAction=""OnTag"" imageMso=""AddTag""/>
          <button id=""btnFeedback"" label=""导入反馈"" size=""normal""
                  onAction=""OnFeedback"" imageMso=""ReviewAcceptChange""/>
        </group>
        <group id=""grpConfig"" label=""配置"">
          <button id=""btnSettings"" label=""插件设置"" size=""large""
                  onAction=""OnSettings"" imageMso=""SourceControlSettings""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}

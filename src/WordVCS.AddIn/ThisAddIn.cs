using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordVCS.AddIn
{
    /// <summary>
    /// WordVCS COM Add-in — loads into Word on startup.
    /// Shows a custom Ribbon tab and a version-control task pane.
    /// </summary>
    [ComVisible(true)]
    [ProgId("WordVCS.Connect")]
    [Guid("B4E1C2D3-A5F6-7890-ABCD-EF1234567890")]
    public sealed class ThisAddIn : IExtensibility, IRibbonExtensibility
    {
        private Word.Application _app;
        private UI.WordVCSTaskPane _wpfPane;
        private TaskPaneManager _mgr;
        private UserControl _host;

        #region IExtensibility (IDTExtensibility2)

        public void OnConnection(object application, ext_ConnectMode mode,
            object addInInst, ref Array custom)
        {
            _app = application as Word.Application;
            if (_app == null) return;

            // WPF needs an Application instance on this thread
            if (System.Windows.Application.Current == null)
                new System.Windows.Application {
                    ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                };

            _app.DocumentOpen += (Word.Document doc) => InitDoc(doc);

            if (_app.ActiveDocument != null) InitDoc(_app.ActiveDocument);
        }

        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom)
        {
            RemovePane();
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID) => RibbonXml.Text;

        #endregion

        #region Ribbon Callbacks (called via XML onAction)

        public void OnShowPanel(IRibbonControl ctl) => ShowPane();

        public void OnCommit(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.ShowCommitDialog(); } catch { }
        }

        public void OnHistory(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.SwitchToTab(0); } catch { }
        }

        public void OnDiff(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.SwitchToTab(2); } catch { }
        }

        public void OnBranch(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.ShowBranchDialog(); } catch { }
        }

        public void OnTag(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.ShowTagDialog(); } catch { }
        }

        public void OnFeedback(IRibbonControl ctl)
        {
            ShowPane();
            try { _mgr?.ShowFeedbackImportDialog(); } catch { }
        }

        public void OnSettings(IRibbonControl ctl)
        {
            try { _mgr?.ShowSettings(); } catch { }
        }

        #endregion

        #region Document / Pane Management

        private void InitDoc(Word.Document doc)
        {
            if (doc == null) return;
            try { if (string.IsNullOrEmpty(doc.FullName)) return; }
            catch { return; }

            RemovePane();

            try
            {
                _wpfPane = new UI.WordVCSTaskPane();
                _host = new UserControl();
                var eh = new ElementHost { Dock = DockStyle.Fill, Child = _wpfPane };
                _host.Controls.Add(eh);
                _host.CreateControl();

                _mgr = new TaskPaneManager(_wpfPane, doc, _app);
                _mgr.Initialize();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordVCS] InitDoc: " + ex.Message);
            }
        }

        public void ShowPane()
        {
            if (_app?.ActiveDocument == null) return;
            if (_mgr == null) InitDoc(_app.ActiveDocument);
        }

        private void RemovePane()
        {
            _host?.Dispose();
            _host = null;
            _wpfPane = null;
            _mgr = null;
        }

        #endregion
    }

    #region COM Interface Definitions

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8A-00A0C90F2744")]
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

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0, ext_cm_Startup = 1,
        ext_cm_External = 2, ext_cm_CommandLine = 3,
        ext_cm_Solution = 4, ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2, ext_dm_SolutionClosed = 3
    }

    #endregion
}

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.UI;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat
{
    [ComVisible(true)]
    [Guid("0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7")]
    [ProgId("OutlookLocalAIChat.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _outlookApplication;
        private object _ctpFactory;
        private object _taskPane;
        private ChatPane _chatPane;

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _outlookApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _outlookApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            CloseTaskPane();
        }

        public void CTPFactoryAvailable(object ctpFactory)
        {
            _ctpFactory = ctpFactory;
        }

        public string GetCustomUI(string ribbonId)
        {
            var tabId = GetTabId(ribbonId);
            if (tabId == null)
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"" + tabId + "\">" +
                "<group id=\"OutlookLocalAIChat.Group\" label=\"AI Chat\">" +
                "<button id=\"OutlookLocalAIChat.Open\" label=\"Mailbox AI Chat\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Mailbox AI Chat\" " +
                "supertip=\"Chat with your mailbox in an Outlook sidebar and open unsent drafts.\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_outlookApplication == null)
                {
                    MessageBox.Show(
                        "Outlook is not ready. Restart Outlook and try again.",
                        "Outlook Local AI Chat",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "Outlook has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "Outlook Local AI Chat",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_taskPane == null)
                {
                    object parentWindow = GetRibbonContext(control);
                    dynamic factory = _ctpFactory;
                    _taskPane = factory.CreateCTP(
                        "OutlookLocalAIChat.ChatPane",
                        "Mailbox AI Chat",
                        parentWindow ?? Type.Missing);

                    dynamic pane = _taskPane;
                    pane.DockPosition = 2;
                    pane.Width = 380;

                    _chatPane = pane.ContentControl as ChatPane ??
                        ChatPane.LastCreated;
                    if (_chatPane == null)
                    {
                        throw new InvalidOperationException(
                            "Outlook created the sidebar but its chat control was unavailable.");
                    }

                    _chatPane.Initialize(_outlookApplication);
                    pane.Visible = true;
                }
                else
                {
                    _chatPane?.RefreshSelectedMessage();
                    dynamic pane = _taskPane;
                    pane.Visible = true;
                }
            }
            catch (Exception exception)
            {
                Log.Error("OnOpenChat", exception);
                MessageBox.Show(
                    DiagnosticDetails.ForException(
                        exception,
                        "SIDEBAR_OPEN_FAILED"),
                    "Outlook Local AI Chat",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string GetTabId(string ribbonId)
        {
            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Explorer",
                StringComparison.Ordinal))
            {
                return "TabMail";
            }

            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Mail.Read",
                StringComparison.Ordinal))
            {
                return "TabReadMessage";
            }

            return null;
        }

        private static object GetRibbonContext(object control)
        {
            try
            {
                dynamic ribbonControl = control;
                return ribbonControl?.Context;
            }
            catch
            {
                return null;
            }
        }

        private void CloseTaskPane()
        {
            try
            {
                _chatPane?.Shutdown();
                if (_taskPane != null)
                {
                    dynamic pane = _taskPane;
                    pane.Visible = false;
                }
            }
            catch (Exception exception)
            {
                Log.Error("CloseTaskPane", exception);
            }

            Release(_taskPane);
            Release(_ctpFactory);
            _taskPane = null;
            _ctpFactory = null;
            _chatPane = null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}

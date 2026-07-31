/*
THESIS: A restrained Outlook sidebar makes mailbox retrieval and one linked,
human-reviewed draft visible without granting send capability.
OWN-WORLD: Windows white and cool-gray surfaces, Outlook blue only for direct actions,
square native fields, one mailbox scope strip, and plain text throughout.
STORY: Ask the mailbox, observe what context was loaded, then deliberately open an
unsent draft.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.UI
{
    [ComVisible(true)]
    [Guid("14D24FA1-4342-442F-B68B-B68D7372794C")]
    [ProgId("OutlookLocalAIChat.ChatPane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ChatPane : UserControl
    {
        private static Color OutlookBlue
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(0, 95, 184);
            }
        }

        private static Color TextPrimary
        {
            get { return SystemColors.WindowText; }
        }

        private static Color TextSecondary
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(80, 80, 80);
            }
        }

        private static Color ErrorText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.HotTrack
                    : Color.FromArgb(163, 38, 38);
            }
        }

        private static Color SurfaceMuted
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Control
                    : Color.FromArgb(244, 246, 248);
            }
        }

        private readonly SettingsStore _settingsStore =
            new SettingsStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly List<ChatTurn> _history =
            new List<ChatTurn>();
        private readonly RichTextBox _transcript =
            new RichTextBox();
        private readonly TextBox _composer = new TextBox();
        private readonly CheckBox _allowOneDraft =
            new CheckBox();
        private readonly Label _scopeTitle = new Label();
        private readonly Label _scopeMeta = new Label();
        private readonly Label _modelMeta = new Label();
        private readonly Label _draftState = new Label();
        private readonly Label _status = new Label();
        private readonly Button _send = new Button();
        private Button _refresh;
        private Button _newChat;
        private Button _settingsButton;

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private DraftToolHost _draftTools;
        private CancellationTokenSource _requestCancellation;
        private bool _busy;
        private bool _shutdown;

        public ChatPane()
        {
            LastCreated = this;
            _settings = _settingsStore.Load();

            Dock = DockStyle.Fill;
            BackColor = SystemColors.Window;
            ForeColor = TextPrimary;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            MinimumSize = new Size(300, 480);
            BuildLayout();
            UpdateModelMeta();
            ShowWelcome();
        }

        internal static ChatPane LastCreated { get; private set; }

        internal void Initialize(object outlookApplication)
        {
            if (_outlookApplication != null)
            {
                return;
            }

            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
            _draftTools = new DraftToolHost(
                _outlookApplication);
            RefreshSelectedMessage();
            UpdateDraftState();
            _composer.Focus();
        }

        public void RefreshSelectedMessage()
        {
            if (_outlookApplication == null)
            {
                SetScopeUnavailable(
                    "Outlook is still initializing.");
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Wait for the request to finish before refreshing the selection.",
                    true);
                return;
            }

            try
            {
                SetSelectedMessage(
                    new MessageReader(_outlookApplication)
                        .CaptureCurrent());
                SetStatus(
                    "The model can search and read bounded context from Inbox and Sent Items.",
                    false);
            }
            catch (Exception exception)
            {
                _selectedMessage = null;
                SetScopeUnavailable(
                    "No selected email. Mailbox search is still available.");
                SetStatus(
                    "Ask about your mailbox, or select an email and refresh the selection.",
                    false);
                Log.Error("CaptureCurrent", exception);
            }
        }

        public void UseRibbonSelection(object selection)
        {
            if (_outlookApplication == null || selection == null)
            {
                RefreshSelectedMessage();
                return;
            }

            try
            {
                SetSelectedMessage(
                    new MessageReader(_outlookApplication)
                        .CaptureSelection(selection));
                SetStatus(
                    "Selected email added. Ask about it or let the model load related mailbox context.",
                    false);
            }
            catch (Exception exception)
            {
                Log.Error("CaptureRibbonSelection", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EMAIL_SELECTION_FAILED");
                SetStatus(FirstLine(details), true);
            }
        }

        internal void Shutdown()
        {
            if (_shutdown)
            {
                return;
            }

            _shutdown = true;
            _requestCancellation?.Cancel();
            _requestCancellation?.Dispose();
            _requestCancellation = null;
            _client.Dispose();
            _draftTools?.Dispose();
            _draftTools = null;
            _outlookApplication = null;
            if (ReferenceEquals(LastCreated, this))
            {
                LastCreated = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Shutdown();
            }

            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(0)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 92));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 154));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildTranscript(), 0, 2);
            root.Controls.Add(BuildComposer(), 0, 3);
            root.Controls.Add(BuildStatusArea(), 0, 4);
            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted,
                Padding = new Padding(14, 10, 14, 8),
                ColumnCount = 1,
                RowCount = 3
            };
            header.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 28));
            header.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 22));
            header.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 22));

            _scopeTitle.AutoEllipsis = true;
            _scopeTitle.Dock = DockStyle.Fill;
            _scopeTitle.Font = new Font(
                Font.FontFamily,
                Font.Size + 2F,
                FontStyle.Bold);
            _scopeTitle.ForeColor = TextPrimary;
            _scopeTitle.Text = "Inbox Cove";

            _scopeMeta.AutoEllipsis = true;
            _scopeMeta.Dock = DockStyle.Fill;
            _scopeMeta.ForeColor = TextSecondary;
            _scopeMeta.Text =
                "No selected email. Mailbox search is available.";

            _modelMeta.AutoEllipsis = true;
            _modelMeta.Dock = DockStyle.Fill;
            _modelMeta.ForeColor = TextSecondary;
            _modelMeta.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Regular);
            _modelMeta.AccessibleName =
                "Active AI model and safety boundary";

            header.Controls.Add(_scopeTitle, 0, 0);
            header.Controls.Add(_scopeMeta, 0, 1);
            header.Controls.Add(_modelMeta, 0, 2);
            return header;
        }

        private Control BuildToolbar()
        {
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 4, 8, 2),
                BackColor = SystemColors.Window
            };

            _refresh = MakeLinkButton("Refresh selection", 116);
            _refresh.Click +=
                (sender, args) => RefreshSelectedMessage();
            _newChat = MakeLinkButton("New chat", 74);
            _newChat.Click += NewChatClick;
            _settingsButton = MakeLinkButton("Settings", 72);
            _settingsButton.Click += SettingsClick;

            toolbar.Controls.Add(_refresh);
            toolbar.Controls.Add(_newChat);
            toolbar.Controls.Add(_settingsButton);
            return toolbar;
        }

        private Control BuildTranscript()
        {
            _transcript.Dock = DockStyle.Fill;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = SystemColors.Window;
            _transcript.ForeColor = TextPrimary;
            _transcript.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _transcript.ReadOnly = true;
            _transcript.DetectUrls = false;
            _transcript.HideSelection = false;
            _transcript.ScrollBars =
                RichTextBoxScrollBars.Vertical;
            _transcript.AccessibleName =
                "Inbox Cove conversation";
            _transcript.AccessibleDescription =
                "Plain-text mailbox conversation and context-loading ledger.";

            var frame = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14, 8, 14, 8),
                BackColor = SystemColors.Window
            };
            frame.Controls.Add(_transcript);
            return frame;
        }

        private Control BuildComposer()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(14, 8, 14, 8),
                BackColor = SurfaceMuted
            };
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 96));
            panel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 22));

            _composer.Dock = DockStyle.Fill;
            _composer.Multiline = true;
            _composer.AcceptsReturn = true;
            _composer.ScrollBars = ScrollBars.Vertical;
            _composer.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _composer.BorderStyle = BorderStyle.FixedSingle;
            _composer.MaxLength =
                TextBoundary.MaxUserPromptCharacters;
            _composer.AccessibleName = "Message to AI";
            _composer.AccessibleDescription =
                "Ask about the mailbox or request draft text. Control Enter sends.";
            _composer.KeyDown += ComposerKeyDown;

            ConfigurePrimaryButton(_send, "Send to AI");
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(8, 0, 0, 0);
            _send.Click += SendClick;

            _allowOneDraft.AutoSize = true;
            _allowOneDraft.Text =
                "Allow one unsent draft for this request";
            _allowOneDraft.ForeColor = TextPrimary;
            _allowOneDraft.Padding = new Padding(0, 3, 0, 0);
            _allowOneDraft.AccessibleName =
                "Allow one unsent draft";
            _allowOneDraft.AccessibleDescription =
                "Adds one create-draft capability to the next AI request. " +
                "The permission resets after Send and cannot send email.";
            _allowOneDraft.CheckedChanged +=
                DraftAuthorizationChanged;

            _draftState.AutoSize = true;
            _draftState.Text =
                "Draft linked. Feedback updates it automatically.";
            _draftState.ForeColor = OutlookBlue;
            _draftState.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Bold);
            _draftState.Padding = new Padding(0, 3, 0, 0);
            _draftState.Visible = false;
            _draftState.AccessibleName = "Linked draft status";

            var hint = new Label
            {
                AutoSize = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text =
                    "Ctrl+Enter sends.",
                Padding = new Padding(0, 3, 0, 0)
            };

            panel.Controls.Add(_composer, 0, 0);
            panel.Controls.Add(_send, 1, 0);
            var draftMode = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted
            };
            draftMode.Controls.Add(_allowOneDraft);
            draftMode.Controls.Add(_draftState);
            panel.Controls.Add(draftMode, 0, 1);
            panel.SetColumnSpan(draftMode, 2);
            panel.Controls.Add(hint, 0, 2);
            panel.SetColumnSpan(hint, 2);
            return panel;
        }

        private Control BuildStatusArea()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted,
                Padding = new Padding(14, 8, 14, 8)
            };
            _status.Dock = DockStyle.Fill;
            _status.AutoEllipsis = true;
            _status.ForeColor = TextSecondary;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AccessibleName = "Chat status";
            _status.AccessibleRole = AccessibleRole.StatusBar;
            _status.Text =
                "Mailbox reads are bounded. Draft creation requires one-shot authorization.";
            panel.Controls.Add(_status);
            return panel;
        }

        private async void SendClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                _requestCancellation?.Cancel();
                return;
            }

            var prompt = TextBoundary.PlainText(
                _composer.Text,
                TextBoundary.MaxUserPromptCharacters);
            if (prompt.Length == 0)
            {
                SetStatus(
                    "Type a mailbox question or drafting instruction first.",
                    true);
                return;
            }

            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing.",
                    true);
                return;
            }

            if (!_settings.IsConfigured)
            {
                OpenSettings();
                if (!_settings.IsConfigured)
                {
                    return;
                }
            }

            var requestSelectedMessage = _selectedMessage;
            var hasLinkedDraft =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    _allowOneDraft.Checked &&
                    !hasLinkedDraft,
                    hasLinkedDraft);
            _allowOneDraft.Checked = false;
            var transcriptStart = _transcript.TextLength;
            AppendTurn("You", prompt, OutlookBlue);
            _composer.Clear();
            SetBusy(true);
            _requestCancellation =
                new CancellationTokenSource();

            try
            {
                var response = await CompleteMailboxChatAsync(
                    requestSelectedMessage,
                    prompt,
                    draftAuthorization,
                    _requestCancellation.Token);

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                AppendTurn(
                    "Assistant",
                    response,
                    TextPrimary);
                if (draftAuthorization.IsCreated)
                {
                    SetStatus(
                        "One unsent draft is open and linked. Feedback now updates it automatically.",
                        false);
                }
                else if (draftAuthorization.IsUpdated)
                {
                    SetStatus(
                        "The linked unsent draft was updated in Outlook.",
                        false);
                }
                else if (draftAuthorization.IsConsumed)
                {
                    SetStatus(
                        draftAuthorization.CanUpdate
                            ? "The linked draft update did not complete. This request cannot retry it."
                            : "Draft creation did not complete. The one-shot permission is consumed.",
                        true);
                }
                else if (draftAuthorization.CanCreate)
                {
                    SetStatus(
                        "Response received without creating a draft. " +
                        "The one-shot permission expired.",
                        false);
                }
                else
                {
                    SetStatus(
                        hasLinkedDraft
                            ? "Response received. The linked draft was unchanged."
                            : "Response received. Draft creation was not authorized.",
                        false);
                }
            }
            catch (OperationCanceledException)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                SetStatus("Request cancelled. Your prompt was restored.", false);
            }
            catch (Exception exception)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "AI_REQUEST_FAILED");
                AppendError(details);
                SetStatus(
                    FirstLine(details),
                    true);
                Log.Error("CompleteMailboxChat", exception);
            }
            finally
            {
                _requestCancellation?.Dispose();
                _requestCancellation = null;
                SetBusy(false);
                UpdateDraftState();
            }
        }

        private async Task<string> CompleteMailboxChatAsync(
            MessageSnapshot selectedMessage,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            CancellationToken cancellationToken)
        {
            var request = ChatRequestFactory.Create(
                _settings.Model,
                selectedMessage,
                _history,
                prompt,
                draftAuthorization.CanCreate,
                _draftTools?.ActiveDraft);
            var mailboxTools = new MailboxToolHost(
                _outlookApplication,
                selectedMessage);
            for (var round = 0;
                 round <= TextBoundary.MaxToolRounds;
                 round++)
            {
                var response = await _client.CompleteAsync(
                    _settings,
                    request,
                    cancellationToken);
                var toolCalls = response.tool_calls;
                if (toolCalls == null || toolCalls.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(response.content))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The model stopped without returning text.");
                    }

                    return response.content;
                }

                if (round == TextBoundary.MaxToolRounds)
                {
                    throw new AiEndpointException(
                        "TOOL_ROUND_LIMIT",
                        "The model exceeded the maximum number of bounded tool rounds.");
                }

                if (toolCalls.Count >
                    TextBoundary.MaxToolCallsPerRound)
                {
                    throw new AiEndpointException(
                        "TOOL_CALL_LIMIT",
                        "The model requested too many tools in one round.");
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var isDraftCall =
                        DraftToolCatalog.IsDraftTool(
                            toolCall?.function?.name);
                    var result = isDraftCall
                        ? _draftTools.Execute(
                            toolCall,
                            selectedMessage,
                            draftAuthorization,
                            toolCalls.Count == 1)
                        : mailboxTools.Execute(toolCall);
                    results.Add(result);
                    if (isDraftCall)
                    {
                        AppendDraftAction(
                            result.StatusText);
                    }
                    else
                    {
                        AppendContext(
                            result.StatusText);
                    }
                    SetStatus(result.StatusText, false);
                }

                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    results);
            }

            throw new AiEndpointException(
                "TOOL_ROUND_LIMIT",
                "The model did not finish after bounded tool use.");
        }

        private void NewChatClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            _history.Clear();
            _draftTools?.Dispose();
            _draftTools = _outlookApplication == null
                ? null
                : new DraftToolHost(_outlookApplication);
            _allowOneDraft.Checked = false;
            _transcript.Clear();
            ShowWelcome();
            UpdateDraftState();
            SetStatus(
                "New mailbox chat started. No previous context is retained.",
                false);
            _composer.Focus();
        }

        private void SettingsClick(
            object sender,
            EventArgs eventArgs)
        {
            OpenSettings();
        }

        private void OpenSettings()
        {
            using (var settingsWindow =
                new SettingsWindow(_settingsStore, _settings))
            {
                if (settingsWindow.ShowDialog(this) ==
                    DialogResult.OK)
                {
                    _settings =
                        settingsWindow.SavedSettings;
                    UpdateModelMeta();
                    SetStatus(
                        "AI endpoint settings saved for " +
                        _settings.Model + ".",
                        false);
                }
            }
        }

        private void AppendContext(string text)
        {
            AppendStyledBlock(
                "Context",
                text,
                TextSecondary,
                FontStyle.Italic);
        }

        private void AppendDraftAction(string text)
        {
            AppendStyledBlock(
                "Draft",
                text,
                OutlookBlue,
                FontStyle.Regular);
        }

        private void AppendError(string text)
        {
            AppendStyledBlock(
                "Error",
                text,
                ErrorText,
                FontStyle.Regular);
        }

        private void AppendStyledBlock(
            string label,
            string text,
            Color color,
            FontStyle bodyStyle)
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            _transcript.SelectionColor = color;
            _transcript.AppendText(label + Environment.NewLine);
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                bodyStyle);
            _transcript.SelectionColor = color;
            _transcript.AppendText(
                TextBoundary.PlainText(text, 2400) +
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendTurn(
            string speaker,
            string text,
            Color headingColor)
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            _transcript.SelectionColor = headingColor;
            _transcript.AppendText(
                speaker + Environment.NewLine);
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _transcript.SelectionColor = TextPrimary;
            _transcript.AppendText(
                TextBoundary.PlainText(
                    text,
                    speaker == "You"
                        ? TextBoundary.MaxUserPromptCharacters
                        : TextBoundary.MaxAssistantCharacters) +
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void ScrollTranscript()
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.ScrollToCaret();
        }

        private void RestoreFailedPrompt(
            string prompt,
            int transcriptStart)
        {
            if (_transcript.TextLength > transcriptStart)
            {
                _transcript.Select(
                    transcriptStart,
                    _transcript.TextLength -
                    transcriptStart);
                _transcript.SelectedText =
                    string.Empty;
            }

            _composer.Text = prompt;
            _composer.SelectionStart =
                _composer.TextLength;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            _send.Text = busy ? "Cancel" : "Send to AI";
            _composer.Enabled = !busy;
            _allowOneDraft.Enabled = !busy;
            _refresh.Enabled = !busy;
            _newChat.Enabled = !busy;
            _settingsButton.Enabled = !busy;
            if (busy)
            {
                SetStatus(
                    "Waiting for the AI endpoint. It may request mailbox context.",
                    false);
            }
        }

        private void UpdateDraftState()
        {
            var linked =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            _allowOneDraft.Visible = !linked;
            _allowOneDraft.Enabled = !linked && !_busy;
            _draftState.Visible = linked;
        }

        private void SetStatus(string text, bool error)
        {
            _status.Text =
                TextBoundary.PlainText(text, 600);
            _status.ForeColor =
                error ? ErrorText : TextSecondary;
        }

        private void SetScopeUnavailable(string text)
        {
            _scopeTitle.Text = "Inbox Cove";
            _scopeMeta.Text = text;
        }

        private void SetSelectedMessage(MessageSnapshot message)
        {
            _selectedMessage = message ??
                throw new ArgumentNullException(nameof(message));
            _scopeTitle.Text = "Inbox Cove";
            var displaySubject = SubjectDisplay.Clean(
                _selectedMessage.Subject);
            _scopeMeta.Text =
                "Selected: " +
                (string.IsNullOrWhiteSpace(displaySubject)
                    ? "(No subject)"
                    : displaySubject);
        }

        private void UpdateModelMeta()
        {
            var model = _settings?.Model ?? string.Empty;
            _modelMeta.Text = "Model: " +
                (model.Length > 0
                    ? model
                    : "not configured");
        }

        private void ShowWelcome()
        {
            AppendStyledBlock(
                "Ready",
                "Ask across Inbox and Sent Items. The model decides which bounded " +
                "messages to load, and every context operation appears here.\n\n" +
                "Try:\n" +
                "- Summarize what needs a reply this week.\n" +
                "- Find decisions about a project or topic.\n" +
                "- Allow one draft, then ask to open a concise reply.\n" +
                "- Once it opens, ask to shorten it or bold an exact section.",
                TextSecondary,
                FontStyle.Regular);
            _transcript.SelectionStart = 0;
            _transcript.ScrollToCaret();
        }

        private void DraftAuthorizationChanged(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            SetStatus(
                _allowOneDraft.Checked
                    ? "One unsent draft is authorized for the next request only."
                    : "Draft creation is not authorized for the next request.",
                false);
        }

        private void ComposerKeyDown(
            object sender,
            KeyEventArgs eventArgs)
        {
            if (eventArgs.Control &&
                eventArgs.KeyCode == Keys.Enter)
            {
                eventArgs.SuppressKeyPress = true;
                SendClick(_send, EventArgs.Empty);
            }
        }

        private static string FirstLine(string value)
        {
            var text = value ?? string.Empty;
            var index = text.IndexOfAny(
                new[] { '\r', '\n' });
            return index >= 0
                ? text.Substring(0, index)
                : text;
        }

        private static Button MakeLinkButton(
            string text,
            int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Window,
                ForeColor = OutlookBlue,
                UseVisualStyleBackColor = false,
                Margin = new Padding(0, 0, 4, 0),
                AccessibleName = text
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static void ConfigurePrimaryButton(
            Button button,
            string text)
        {
            button.Text = text;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = OutlookBlue;
            button.ForeColor =
                SystemColors.HighlightText;
            button.UseVisualStyleBackColor = false;
            button.FlatAppearance.BorderColor =
                OutlookBlue;
            button.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            button.AccessibleName = text;
        }

    }
}

/*
THESIS: A restrained Outlook sidebar makes mailbox retrieval visible while keeping
draft creation separate from model-controlled read-only context.
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
        private readonly Label _scopeTitle = new Label();
        private readonly Label _scopeMeta = new Label();
        private readonly Label _modelMeta = new Label();
        private readonly Label _status = new Label();
        private readonly Button _send = new Button();
        private readonly Button _replyDraft = new Button();
        private readonly Button _newDraft = new Button();
        private Button _refresh;
        private Button _newChat;
        private Button _settingsButton;

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private MessageSnapshot _draftSource;
        private string _lastAssistantText = string.Empty;
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
            RefreshSelectedMessage();
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
                _selectedMessage =
                    new MessageReader(_outlookApplication)
                        .CaptureCurrent();
                _scopeTitle.Text = "Mailbox chat";
                _scopeMeta.Text =
                    "Selected: " +
                    (string.IsNullOrWhiteSpace(
                        _selectedMessage.Subject)
                        ? "(No subject)"
                        : _selectedMessage.Subject);
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
                RowCount = 6,
                Padding = new Padding(0)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 92));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildTranscript(), 0, 2);
            root.Controls.Add(BuildComposer(), 0, 3);
            root.Controls.Add(BuildDraftActions(), 0, 4);
            root.Controls.Add(BuildStatusArea(), 0, 5);
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
            _scopeTitle.Text = "Mailbox chat";

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
                "Mailbox AI chat conversation";
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
                RowCount = 2,
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

            var hint = new Label
            {
                AutoSize = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text =
                    "The model chooses read-only mailbox context. Ctrl+Enter sends.",
                Padding = new Padding(0, 3, 0, 0)
            };

            panel.Controls.Add(_composer, 0, 0);
            panel.Controls.Add(_send, 1, 0);
            panel.Controls.Add(hint, 0, 1);
            panel.SetColumnSpan(hint, 2);
            return panel;
        }

        private Control BuildDraftActions()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(14, 3, 14, 3),
                BackColor = SystemColors.Window
            };
            panel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 20));
            panel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            var disclosure = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text =
                    "Drafts use the entire latest assistant response.",
                AccessibleName = "Draft content disclosure"
            };

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 2, 0, 0),
                BackColor = SystemColors.Window
            };

            ConfigureSecondaryButton(
                _replyDraft,
                "Reply draft",
                106);
            _replyDraft.Click += ReplyDraftClick;
            ConfigureSecondaryButton(
                _newDraft,
                "New draft",
                96);
            _newDraft.Click += NewDraftClick;

            actions.Controls.Add(_replyDraft);
            actions.Controls.Add(_newDraft);
            panel.Controls.Add(disclosure, 0, 0);
            panel.Controls.Add(actions, 0, 1);
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
                "Mailbox tools are read-only. Drafts require your click.";
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
                    _requestCancellation.Token);

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                _lastAssistantText = response;
                _draftSource = requestSelectedMessage;
                AppendTurn(
                    "Assistant",
                    response,
                    TextPrimary);
                SetStatus(
                    "Response received. Draft actions still require your click.",
                    false);
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
            }
        }

        private async Task<string> CompleteMailboxChatAsync(
            MessageSnapshot selectedMessage,
            string prompt,
            CancellationToken cancellationToken)
        {
            var request = ChatRequestFactory.Create(
                _settings.Model,
                selectedMessage,
                _history,
                prompt);
            var tools = new MailboxToolHost(
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
                        "MAILBOX_TOOL_ROUND_LIMIT",
                        "The model exceeded the maximum number of mailbox context rounds.");
                }

                if (toolCalls.Count >
                    TextBoundary.MaxToolCallsPerRound)
                {
                    throw new AiEndpointException(
                        "MAILBOX_TOOL_CALL_LIMIT",
                        "The model requested too many mailbox tools in one round.");
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var result = tools.Execute(toolCall);
                    results.Add(result);
                    AppendContext(result.StatusText);
                    SetStatus(result.StatusText, false);
                }

                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    results);
            }

            throw new AiEndpointException(
                "MAILBOX_TOOL_ROUND_LIMIT",
                "The model did not finish after bounded mailbox context retrieval.");
        }

        private void ReplyDraftClick(
            object sender,
            EventArgs eventArgs)
        {
            try
            {
                new DraftService(_outlookApplication)
                    .CreateReplyDraft(
                        _draftSource,
                        _lastAssistantText);
                SetStatus(
                    "Unsent reply draft opened in Outlook for your review.",
                    false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "REPLY_DRAFT_FAILED");
                AppendError(details);
                SetStatus(FirstLine(details), true);
                Log.Error("CreateReplyDraft", exception);
            }
        }

        private void NewDraftClick(
            object sender,
            EventArgs eventArgs)
        {
            try
            {
                new DraftService(_outlookApplication)
                    .CreateNewDraft(_lastAssistantText);
                SetStatus(
                    "Unsent new-message draft opened in Outlook for your review.",
                    false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "NEW_DRAFT_FAILED");
                AppendError(details);
                SetStatus(FirstLine(details), true);
                Log.Error("CreateNewDraft", exception);
            }
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
            _lastAssistantText = string.Empty;
            _draftSource = null;
            _transcript.Clear();
            ShowWelcome();
            UpdateDraftButtons();
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
            _refresh.Enabled = !busy;
            _newChat.Enabled = !busy;
            _settingsButton.Enabled = !busy;
            UpdateDraftButtons();
            if (busy)
            {
                SetStatus(
                    "Waiting for the AI endpoint. It may request mailbox context.",
                    false);
            }
        }

        private void UpdateDraftButtons()
        {
            var enabled =
                !_busy &&
                _lastAssistantText.Length > 0;
            _newDraft.Enabled = enabled;
            _replyDraft.Enabled =
                enabled &&
                _draftSource != null &&
                _draftSource.CanReply;
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
            _scopeTitle.Text = "Mailbox chat";
            _scopeMeta.Text = text;
        }

        private void UpdateModelMeta()
        {
            var configured =
                _settings != null &&
                _settings.IsConfigured;
            var model = _settings?.Model ?? string.Empty;
            _modelMeta.Text = configured
                ? "Model: " + model +
                  "  |  Search and read only"
                : "Setup required  |  Search and read only";
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
                "- Draft a concise reply based on the full conversation.",
                TextSecondary,
                FontStyle.Regular);
            _transcript.SelectionStart = 0;
            _transcript.ScrollToCaret();
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

        private static void ConfigureSecondaryButton(
            Button button,
            string text,
            int width)
        {
            button.Text = text;
            button.Width = width;
            button.Height = 34;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = SystemColors.Window;
            button.ForeColor = TextPrimary;
            button.UseVisualStyleBackColor = false;
            button.FlatAppearance.BorderColor =
                SystemColors.ControlDark;
            button.Margin = new Padding(8, 0, 0, 0);
            button.AccessibleName = text;
        }
    }
}

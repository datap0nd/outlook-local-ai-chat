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
        private const int WorkingSetExpandedHeight = 322;
        private const int WorkingSetCollapsedHeight = 36;

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
        private readonly List<MessageSnapshot> _workingMessages =
            new List<MessageSnapshot>();
        private readonly List<ExternalContextDocument> _externalContext =
            new List<ExternalContextDocument>();
        private readonly RichTextBox _transcript =
            new RichTextBox();
        private readonly TextBox _composer = new TextBox();
        private readonly Label _scopeTitle = new Label();
        private readonly Label _scopeMeta = new Label();
        private readonly Label _modelMeta = new Label();
        private readonly Label _draftState = new Label();
        private readonly Label _status = new Label();
        private readonly Button _send = new Button();
        private readonly Panel _workingSetLayer = new Panel();
        private readonly Label _workingSetHeading = new Label();
        private readonly FlowLayoutPanel _workingSetCards =
            new FlowLayoutPanel();
        private TableLayoutPanel _rootLayout;
        private Button _workingSetToggle;
        private Button _workingSetClear;
        private Button _refresh;
        private Button _addFiles;
        private Button _newChat;
        private Button _settingsButton;

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private DraftToolHost _draftTools;
        private CancellationTokenSource _requestCancellation;
        private bool _busy;
        private bool _shutdown;
        private bool _workingSetExpanded = true;

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
            AllowDrop = true;
            DragEnter += ChatPaneDragEnter;
            DragDrop += ChatPaneDragDrop;
            BuildLayout();
            UpdateModelMeta();
            ShowWelcome();
        }

        internal static ChatPane LastCreated { get; private set; }

        internal void Initialize(
            object outlookApplication,
            bool refreshSelection = true)
        {
            if (_outlookApplication != null)
            {
                return;
            }

            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
            _draftTools = new DraftToolHost(
                _outlookApplication);
            if (refreshSelection)
            {
                RefreshSelectedMessage();
            }
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
                _workingMessages.Clear();
                RefreshContextLayer("External files");
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
            if (_outlookApplication == null)
            {
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Wait for the active request before changing context.",
                    true);
                return;
            }

            try
            {
                var reader = new MessageReader(
                    _outlookApplication);
                IReadOnlyList<MessageSnapshot> messages;
                try
                {
                    messages = selection == null
                        ? reader.CaptureActiveSelectionMany()
                        : reader.CaptureSelectionMany(selection);
                }
                catch when (selection != null)
                {
                    messages = reader.CaptureActiveSelectionMany();
                }

                ApplySelectedMessages(messages);
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

        public void AddActiveSelection()
        {
            UseRibbonSelection(null);
        }

        private void ApplySelectedMessages(
            IReadOnlyList<MessageSnapshot> messages)
        {
            if (messages == null || messages.Count == 0)
            {
                throw new InvalidOperationException(
                    "Select one to five emails in Outlook first.");
            }

            if (messages.Count == 1)
            {
                SetSelectedMessage(messages[0]);
                SetStatus(
                    "Selected email added. Its body is loaded only if the model requests it.",
                    false);
                return;
            }

            SetWorkingMessages(
                messages,
                messages.Count + " emails selected in Outlook");
        }

        private void HandleLocalSearchCommand(
            string prompt,
            LocalSearchCommand command)
        {
            AppendTurn("You", prompt, OutlookBlue);
            _composer.Clear();
            switch (command.Kind)
            {
                case LocalSearchCommandKind.Help:
                    AppendContext(
                        "Use /search <person or topic> to replace the working set with " +
                        "the newest five matching emails from Inbox and Sent Items. " +
                        "Use /search clear to remove the working set.");
                    SetStatus(
                        "Search help shown. No email context changed.",
                        false);
                    return;
                case LocalSearchCommandKind.Clear:
                    _workingMessages.Clear();
                    RefreshContextLayer("External files");
                    _selectedMessage = null;
                    SetScopeUnavailable(
                        "No working set. Use /search or select email in Outlook.");
                    AppendContext(
                        "Working set cleared. No email bodies are loaded.");
                    SetStatus(
                        "Working set cleared.",
                        false);
                    return;
                case LocalSearchCommandKind.Search:
                    SearchWorkingMessages(command.Query);
                    return;
                default:
                    return;
            }
        }

        private void SearchWorkingMessages(string query)
        {
            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing.",
                    true);
                return;
            }

            UseWaitCursor = true;
            SetStatus(
                "Searching Outlook locally for the newest five matches...",
                false);
            try
            {
                var hits = new MailboxContextService(
                    _outlookApplication)
                    .Search(
                        query,
                        "all",
                        3650,
                        MailboxWorkingSet.MaxMessages);
                var messages = new List<MessageSnapshot>();
                foreach (var hit in hits)
                {
                    messages.Add(hit.Message);
                }

                if (messages.Count == 0)
                {
                    AppendContext(
                        "No emails matched '" + query + "'. " +
                        (_workingMessages.Count > 0
                            ? "The previous working set was kept."
                            : "No working set was created.") +
                        " Refine the person or topic and run /search again.");
                    SetStatus(
                        "No matching emails. Refine the query and try /search again.",
                        true);
                    return;
                }

                SetWorkingMessages(
                    messages,
                    "Search: " + query);
            }
            catch (Exception exception)
            {
                Log.Error("LocalMailboxSearch", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "LOCAL_SEARCH_FAILED");
                AppendError(details);
                SetStatus(FirstLine(details), true);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void SetWorkingMessages(
            IEnumerable<MessageSnapshot> messages,
            string source)
        {
            var bounded = MailboxWorkingSet.Normalize(messages);
            _workingMessages.Clear();
            foreach (var message in bounded)
            {
                _workingMessages.Add(message);
            }

            _selectedMessage = null;
            _scopeTitle.Text = "MailAI";
            _scopeMeta.Text =
                "Working set: " +
                _workingMessages.Count +
                " of 5 emails";
            ShowWorkingSetLayer(source, _workingMessages);
            AppendContext(
                TextBoundary.SingleLine(source, 260) +
                ". The five-email context layer is ready. " +
                "Search again to replace it if needed.");
            SetStatus(
                "Working set ready. Run /search again to replace it, or ask MailAI to work on these emails.",
                false);
        }

        private void ShowWorkingSetLayer(
            string source,
            IReadOnlyList<MessageSnapshot> messages)
        {
            RefreshContextLayer(source);
        }

        private void RefreshContextLayer(string source)
        {
            ClearWorkingSetCards();
            var emailCount = _workingMessages.Count;
            var fileCount = _externalContext.Count;
            if (emailCount + fileCount == 0)
            {
                HideWorkingSetLayer();
                return;
            }

            _workingSetHeading.Text =
                "Context - " +
                (emailCount + fileCount) +
                ((emailCount + fileCount) == 1
                    ? " item"
                    : " items");
            _workingSetHeading.AccessibleDescription =
                TextBoundary.SingleLine(source, 260) +
                ". Bounded user-approved email and file context.";
            for (var index = 0; index < emailCount; index++)
            {
                _workingSetCards.Controls.Add(
                    BuildWorkingSetCard(
                        index,
                        _workingMessages[index]));
            }

            for (var index = 0; index < fileCount; index++)
            {
                _workingSetCards.Controls.Add(
                    BuildExternalContextCard(
                        index,
                        _externalContext[index]));
            }

            _workingSetExpanded = true;
            _workingSetCards.Visible = true;
            _workingSetToggle.Text = "Hide";
            _workingSetToggle.AccessibleName =
                "Hide working set";
            _workingSetLayer.Visible = true;
            SetWorkingSetRowHeight(
                WorkingSetExpandedHeight);
            ResizeWorkingSetCards();
        }

        private Control BuildExternalContextCard(
            int index,
            ExternalContextDocument document)
        {
            var card = new Panel
            {
                Height = 50,
                Margin = new Padding(0, 0, 0, 5),
                Padding = new Padding(8, 4, 8, 4),
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                AccessibleName = "Context file: " + document.Name,
                AccessibleDescription =
                    document.Content.Length + " bounded text characters"
            };
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = SystemColors.Window,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 36));
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 52));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 48));
            var badge = new Label
            {
                Dock = DockStyle.Fill,
                Text = "FILE",
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = OutlookBlue,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(7F, Font.Size - 2F),
                    FontStyle.Bold)
            };
            var name = new Label
            {
                Dock = DockStyle.Fill,
                Text = document.Name,
                AutoEllipsis = true,
                ForeColor = TextPrimary,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var metadata = new Label
            {
                Dock = DockStyle.Fill,
                Text = document.Content.Length + " text characters",
                AutoEllipsis = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular)
            };
            grid.Controls.Add(badge, 0, 0);
            grid.SetRowSpan(badge, 2);
            grid.Controls.Add(name, 1, 0);
            grid.Controls.Add(metadata, 1, 1);
            card.Controls.Add(grid);
            return card;
        }

        private Control BuildWorkingSetCard(
            int index,
            MessageSnapshot message)
        {
            var subject = TextBoundary.SingleLine(
                SubjectDisplay.Clean(message.Subject),
                180);
            if (subject.Length == 0)
            {
                subject = "(No subject)";
            }

            var sender = TextBoundary.SingleLine(
                message.Sender,
                120);
            if (sender.Length == 0)
            {
                sender = "Unknown sender";
            }

            var date = message.ReceivedAt?.ToString(
                "yyyy-MM-dd HH:mm") ??
                "Unknown date";
            var card = new Panel
            {
                Height = 50,
                Margin = new Padding(0, 0, 0, 5),
                Padding = new Padding(8, 4, 8, 4),
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                AccessibleName =
                    "Email " +
                    (index + 1) +
                    ": " +
                    subject,
                AccessibleDescription =
                    sender + ", " + date
            };
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = SystemColors.Window,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 28));
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 52));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 48));

            var number = new Label
            {
                Dock = DockStyle.Fill,
                Text = (index + 1).ToString(),
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = OutlookBlue,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var subjectLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = subject,
                AutoEllipsis = true,
                ForeColor = TextPrimary,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var metadata = new Label
            {
                Dock = DockStyle.Fill,
                Text = sender + " | " + date,
                AutoEllipsis = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular)
            };
            grid.Controls.Add(number, 0, 0);
            grid.SetRowSpan(number, 2);
            grid.Controls.Add(subjectLabel, 1, 0);
            grid.Controls.Add(metadata, 1, 1);
            card.Controls.Add(grid);
            return card;
        }

        private void WorkingSetToggleClick(
            object sender,
            EventArgs eventArgs)
        {
            SetWorkingSetExpanded(
                !_workingSetExpanded);
        }

        private void SetWorkingSetExpanded(bool expanded)
        {
            _workingSetExpanded = expanded;
            _workingSetCards.Visible = expanded;
            _workingSetToggle.Text =
                expanded ? "Hide" : "Show";
            _workingSetToggle.AccessibleName =
                (expanded ? "Hide" : "Show") +
                " context";
            SetWorkingSetRowHeight(
                expanded
                    ? WorkingSetExpandedHeight
                    : WorkingSetCollapsedHeight);
        }

        private void ResizeWorkingSetCards()
        {
            var width = Math.Max(
                120,
                _workingSetCards.ClientSize.Width -
                SystemInformation.VerticalScrollBarWidth -
                2);
            foreach (Control card in
                _workingSetCards.Controls)
            {
                card.Width = width;
            }
        }

        private void HideWorkingSetLayer()
        {
            ClearWorkingSetCards();
            _workingSetLayer.Visible = false;
            SetWorkingSetRowHeight(0);
        }

        private void ClearContextClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            _workingMessages.Clear();
            _externalContext.Clear();
            _selectedMessage = null;
            HideWorkingSetLayer();
            SetScopeUnavailable(
                "No context selected. Use /search, Add email, or Add files.");
            SetStatus(
                "Email and external file context cleared.",
                false);
        }

        private void AddFilesClick(
            object sender,
            EventArgs eventArgs)
        {
            using (var dialog = new OpenFileDialog
            {
                Title = "Add bounded text context to MailAI",
                Multiselect = true,
                CheckFileExists = true,
                Filter =
                    "Supported text files|*.txt;*.md;*.csv;*.json;*.xml;*.html;*.htm;*.log;*.yaml;*.yml;*.ini|" +
                    "All files|*.*"
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    AddExternalFiles(dialog.FileNames);
                }
            }
        }

        private void AddExternalFiles(IEnumerable<string> paths)
        {
            try
            {
                var loaded = ExternalContextLoader.LoadFiles(paths);
                var combined = new List<ExternalContextDocument>(
                    _externalContext);
                combined.AddRange(loaded);
                var normalized =
                    ExternalContextDocument.Normalize(combined);
                _externalContext.Clear();
                foreach (var document in normalized)
                {
                    _externalContext.Add(document);
                }

                RefreshContextLayer("External files");
                SetStatus(
                    _externalContext.Count +
                    (_externalContext.Count == 1
                        ? " external context file is ready."
                        : " external context files are ready.") +
                    " File text is bounded and treated as untrusted data.",
                    false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EXTERNAL_CONTEXT_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("AddExternalContext", exception);
            }
        }

        private void ChatPaneDragEnter(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                eventArgs.Effect = DragDropEffects.None;
                return;
            }

            if (HasOutlookDragFormat(eventArgs.Data))
            {
                eventArgs.Effect = DragDropEffects.Link;
                return;
            }

            if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
            {
                eventArgs.Effect = DragDropEffects.Copy;
                return;
            }

            eventArgs.Effect = DragDropEffects.None;
        }

        private void ChatPaneDragDrop(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                return;
            }

            try
            {
                if (HasOutlookDragFormat(eventArgs.Data))
                {
                    AddActiveSelection();
                    return;
                }

                if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var paths = eventArgs.Data.GetData(
                        DataFormats.FileDrop) as string[];
                    AddExternalFiles(paths);
                    return;
                }

            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "CONTEXT_DROP_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("DropContext", exception);
            }
        }

        private static bool HasOutlookDragFormat(IDataObject data)
        {
            foreach (var format in data.GetFormats())
            {
                if (format.Equals(
                        "RenPrivateMessages",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptor",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptorW",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void ClearWorkingSetCards()
        {
            while (_workingSetCards.Controls.Count > 0)
            {
                var card = _workingSetCards.Controls[0];
                _workingSetCards.Controls.RemoveAt(0);
                card.Dispose();
            }
        }

        private void SetWorkingSetRowHeight(float height)
        {
            if (_rootLayout == null ||
                _rootLayout.RowStyles.Count < 3)
            {
                return;
            }

            _rootLayout.RowStyles[2].Height = height;
            _rootLayout.PerformLayout();
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
            _rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(0)
            };
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 92));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 38));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 0));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 154));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 64));

            _rootLayout.Controls.Add(BuildHeader(), 0, 0);
            _rootLayout.Controls.Add(BuildToolbar(), 0, 1);
            _rootLayout.Controls.Add(BuildWorkingSetLayer(), 0, 2);
            _rootLayout.Controls.Add(BuildTranscript(), 0, 3);
            _rootLayout.Controls.Add(BuildComposer(), 0, 4);
            _rootLayout.Controls.Add(BuildStatusArea(), 0, 5);
            Controls.Add(_rootLayout);
        }

        private Control BuildWorkingSetLayer()
        {
            _workingSetLayer.Dock = DockStyle.Fill;
            _workingSetLayer.BackColor = SurfaceMuted;
            _workingSetLayer.Padding = new Padding(14, 0, 14, 8);
            _workingSetLayer.Visible = false;
            _workingSetLayer.AccessibleName =
                "Selected email working set";

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = WorkingSetCollapsedHeight,
                BackColor = SurfaceMuted
            };
            _workingSetHeading.Dock = DockStyle.Fill;
            _workingSetHeading.TextAlign =
                ContentAlignment.MiddleLeft;
            _workingSetHeading.Font = new Font(
                Font.FontFamily,
                Font.Size,
                FontStyle.Bold);
            _workingSetHeading.ForeColor = TextPrimary;

            _workingSetToggle = MakeLinkButton("Hide", 52);
            _workingSetToggle.Dock = DockStyle.Right;
            _workingSetToggle.Click += WorkingSetToggleClick;
            _workingSetClear = MakeLinkButton("Clear", 52);
            _workingSetClear.Dock = DockStyle.Right;
            _workingSetClear.Click += ClearContextClick;
            header.Controls.Add(_workingSetHeading);
            header.Controls.Add(_workingSetClear);
            header.Controls.Add(_workingSetToggle);

            _workingSetCards.Dock = DockStyle.Fill;
            _workingSetCards.AutoScroll = true;
            _workingSetCards.FlowDirection =
                FlowDirection.TopDown;
            _workingSetCards.WrapContents = false;
            _workingSetCards.Padding = new Padding(0, 0, 0, 4);
            _workingSetCards.BackColor = SurfaceMuted;
            _workingSetCards.Resize +=
                (sender, args) => ResizeWorkingSetCards();

            _workingSetLayer.Controls.Add(_workingSetCards);
            _workingSetLayer.Controls.Add(header);
            return _workingSetLayer;
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
            _scopeTitle.Text = "MailAI";

            _scopeMeta.AutoEllipsis = true;
            _scopeMeta.Dock = DockStyle.Fill;
            _scopeMeta.ForeColor = TextSecondary;
            _scopeMeta.Text =
                "No context selected. Use /search or select emails.";

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

            _refresh = MakeLinkButton("Add email", 74);
            _refresh.Click +=
                (sender, args) => AddActiveSelection();
            _addFiles = MakeLinkButton("Add files", 68);
            _addFiles.Click += AddFilesClick;
            _newChat = MakeLinkButton("New", 44);
            _newChat.Click += NewChatClick;
            _settingsButton = MakeLinkButton("Settings", 64);
            _settingsButton.Click += SettingsClick;

            toolbar.Controls.Add(_refresh);
            toolbar.Controls.Add(_addFiles);
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
                "MailAI conversation";
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
                "Ask about the mailbox or request draft text. Control Enter submits the prompt.";
            _composer.KeyDown += ComposerKeyDown;

            ConfigurePrimaryButton(_send, "Send to AI");
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(8, 0, 0, 0);
            _send.Click += SendClick;

            _draftState.AutoSize = false;
            _draftState.AutoEllipsis = true;
            _draftState.Dock = DockStyle.Fill;
            _draftState.Text =
                "Say 'create a draft' to open one. MailAI cannot send.";
            _draftState.ForeColor = TextSecondary;
            _draftState.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Regular);
            _draftState.Padding = new Padding(0, 3, 0, 0);
            _draftState.Visible = true;
            _draftState.AccessibleName = "Draft safety status";

            var hint = new Label
            {
                AutoSize = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text =
                    "Ctrl+Enter submits prompt.",
                Padding = new Padding(0, 3, 0, 0)
            };

            panel.Controls.Add(_composer, 0, 0);
            panel.Controls.Add(_send, 1, 0);
            var draftMode = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted
            };
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
                "Mailbox reads are capped at five emails. MailAI cannot send.";
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

            var localCommand = LocalSearchCommand.Parse(prompt);
            if (localCommand.Kind != LocalSearchCommandKind.None)
            {
                HandleLocalSearchCommand(prompt, localCommand);
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
            var requestWorkingMessages =
                new List<MessageSnapshot>(_workingMessages);
            var requestExternalContext =
                new List<ExternalContextDocument>(_externalContext);
            var hasLinkedDraft =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    !hasLinkedDraft &&
                    DraftIntentPolicy.AllowsCreate(prompt),
                    hasLinkedDraft &&
                    DraftIntentPolicy.AllowsUpdate(prompt));
            var transcriptStart = _transcript.TextLength;
            AppendTurn("You", prompt, OutlookBlue);
            if (_workingMessages.Count > 0 &&
                _workingSetExpanded)
            {
                SetWorkingSetExpanded(false);
            }
            _composer.Clear();
            SetBusy(true);
            _requestCancellation =
                new CancellationTokenSource();

            try
            {
                var response = await CompleteMailboxChatAsync(
                    requestSelectedMessage,
                    requestWorkingMessages,
                    requestExternalContext,
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
                            : "Draft creation did not complete. This request cannot retry it.",
                        true);
                }
                else if (draftAuthorization.CanCreate)
                {
                    SetStatus(
                        "Draft request recognized, but no Outlook draft was created.",
                        false);
                }
                else
                {
                    SetStatus(
                        hasLinkedDraft
                            ? "Response received. The linked draft was unchanged."
                            : "Response received. Say 'create a draft' when you want one opened.",
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
            IReadOnlyList<MessageSnapshot> workingMessages,
            IReadOnlyList<ExternalContextDocument> externalContext,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            CancellationToken cancellationToken)
        {
            var activeDraft = draftAuthorization.CanUpdate
                ? _draftTools?.ActiveDraft
                : null;
            var request = ChatRequestFactory.Create(
                _settings.Model,
                selectedMessage,
                _history,
                prompt,
                draftAuthorization.CanCreate,
                activeDraft,
                draftAuthorization.CanUpdate,
                workingMessages,
                externalContext,
                _settings.UseToneProfile
                    ? _settings.ToneProfile
                    : null);
            var mailboxTools = new MailboxToolHost(
                _outlookApplication,
                selectedMessage,
                workingMessages);
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
                            mailboxTools.ResolveHandle,
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
            _workingMessages.Clear();
            _externalContext.Clear();
            HideWorkingSetLayer();
            _draftTools?.Dispose();
            _draftTools = _outlookApplication == null
                ? null
                : new DraftToolHost(_outlookApplication);
            _transcript.Clear();
            RefreshSelectedMessage();
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
                new SettingsWindow(
                    _settingsStore,
                    _settings,
                    _outlookApplication))
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
            if (speaker == "Assistant")
            {
                AppendFormattedAssistantText(text);
            }
            else
            {
                SetTranscriptBodyStyle(FontStyle.Regular);
                _transcript.AppendText(
                    TextBoundary.PlainText(
                        text,
                        TextBoundary.MaxUserPromptCharacters));
            }

            SetTranscriptBodyStyle(FontStyle.Regular);
            _transcript.AppendText(
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendFormattedAssistantText(string text)
        {
            var formatted = SafeModelText.Format(
                text,
                TextBoundary.MaxAssistantCharacters);
            var position = 0;
            foreach (var range in formatted.BoldRanges)
            {
                if (range.Start > position)
                {
                    SetTranscriptBodyStyle(
                        FontStyle.Regular);
                    _transcript.AppendText(
                        formatted.PlainText.Substring(
                            position,
                            range.Start - position));
                }

                SetTranscriptBodyStyle(FontStyle.Bold);
                _transcript.AppendText(
                    formatted.PlainText.Substring(
                        range.Start,
                        range.Length));
                position = range.Start + range.Length;
            }

            if (position < formatted.PlainText.Length)
            {
                SetTranscriptBodyStyle(FontStyle.Regular);
                _transcript.AppendText(
                    formatted.PlainText.Substring(position));
            }
        }

        private void SetTranscriptBodyStyle(FontStyle style)
        {
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                style);
            _transcript.SelectionColor = TextPrimary;
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
            _addFiles.Enabled = !busy;
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
            _draftState.Text = linked
                ? "One draft linked. Revision requests update this draft only."
                : "Say 'create a draft' to open one. MailAI cannot send.";
            _draftState.ForeColor = linked
                ? OutlookBlue
                : TextSecondary;
            var style = linked
                ? FontStyle.Bold
                : FontStyle.Regular;
            if (_draftState.Font.Style != style)
            {
                var previousFont = _draftState.Font;
                _draftState.Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    style);
                previousFont.Dispose();
            }
            _draftState.Visible = true;
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
            _scopeTitle.Text = "MailAI";
            _scopeMeta.Text = text;
        }

        private void SetSelectedMessage(MessageSnapshot message)
        {
            _workingMessages.Clear();
            RefreshContextLayer("External files");
            _selectedMessage = message ??
                throw new ArgumentNullException(nameof(message));
            _scopeTitle.Text = "MailAI";
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
                "Ask across Inbox and Sent Items. MailAI loads no more than five " +
                "email bodies. Add up to three bounded text files as external context.\n\n" +
                "Try:\n" +
                "- /search person or topic\n" +
                "- Ctrl+click up to five emails, then click Add email or right-click Send to MailAI.\n" +
                "- Drag selected Outlook emails onto MailAI, or drag supported text files here.\n" +
                "- Ask MailAI to summarize or compare the working set.\n" +
                "- Find a message and create a concise reply draft.\n" +
                "- Once it opens, ask to shorten it or bold an exact section.",
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

    }
}

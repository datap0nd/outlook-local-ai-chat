using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.UI
{
    public sealed class SettingsWindow : Form
    {
        private readonly TextBox _endpoint = new TextBox();
        private readonly ComboBox _model = new ComboBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly CheckBox _allowInsecureHttp =
            new CheckBox();
        private readonly Label _transportWarning = new Label();
        private readonly Label _modelGuidance = new Label();
        private readonly Label _testStatus = new Label();
        private readonly Label _error = new Label();
        private readonly Button _checkEndpoint =
            MakeButton("Check endpoint", false, 128);
        private readonly Button _useRecommended =
            MakeButton("Use recommended", false, 134);
        private readonly Button _save =
            MakeButton("Save", true, 96);
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly SettingsStore _store;
        private CancellationTokenSource _checkCancellation;
        private bool _checking;

        public SettingsWindow(SettingsStore store, AppSettings current)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));

            Text = "Mailbox AI settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(620, 590);
            MinimumSize = new Size(560, 570);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 14,
                Padding = new Padding(24, 20, 24, 20)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 1));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));

            ConfigureField(
                _endpoint,
                "AI endpoint",
                "HTTPS endpoint, loopback HTTP, or explicitly allowed remote HTTP.");
            ConfigureModelField();
            ConfigureField(
                _apiKey,
                "API key",
                "Encrypted for the current Windows user.");
            _apiKey.UseSystemPasswordChar = true;

            layout.Controls.Add(FieldLabel("Endpoint or base URL"), 0, 0);
            layout.Controls.Add(_endpoint, 0, 1);
            layout.Controls.Add(FieldLabel("Model"), 0, 2);
            layout.Controls.Add(BuildModelRow(), 0, 3);

            ConfigureSupportingLabel(_modelGuidance);
            _modelGuidance.AccessibleName = "Model recommendation";
            layout.Controls.Add(_modelGuidance, 0, 4);

            layout.Controls.Add(FieldLabel("API key"), 0, 5);
            layout.Controls.Add(_apiKey, 0, 6);

            _allowInsecureHttp.AutoSize = true;
            _allowInsecureHttp.Text =
                "Allow insecure HTTP for non-local endpoints";
            _allowInsecureHttp.AccessibleName =
                "Allow insecure HTTP";
            _allowInsecureHttp.AccessibleDescription =
                "Allows the API key, prompts, and email context to be sent " +
                "without transport encryption.";
            _allowInsecureHttp.CheckedChanged +=
                InsecureHttpChanged;
            layout.Controls.Add(_allowInsecureHttp, 0, 7);

            ConfigureSupportingLabel(_transportWarning);
            _transportWarning.AccessibleName =
                "HTTP transport warning";
            _transportWarning.AccessibleRole =
                AccessibleRole.Alert;
            layout.Controls.Add(_transportWarning, 0, 8);

            var hint = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(560, 0),
                ForeColor = SecondaryText,
                Text =
                    "The model receives your prompt, recent conversation, and only the " +
                    "bounded mailbox context it requests. The local host exposes search " +
                    "and read tools only. Sending and mailbox changes are not available.",
                AccessibleName = "Endpoint data disclosure"
            };
            layout.Controls.Add(hint, 0, 9);

            ConfigureSupportingLabel(_testStatus);
            _testStatus.AccessibleName = "Endpoint check status";
            _testStatus.AccessibleRole = AccessibleRole.StatusBar;
            _testStatus.Text =
                "Check the endpoint before saving to verify authentication, model " +
                "discovery, and mailbox tool-call compatibility.";
            layout.Controls.Add(_testStatus, 0, 10);

            _error.AutoSize = true;
            _error.MaximumSize = new Size(560, 0);
            _error.ForeColor = ErrorText;
            _error.AccessibleName = "Settings error";
            _error.AccessibleRole = AccessibleRole.Alert;
            layout.Controls.Add(_error, 0, 11);

            layout.Controls.Add(BuildButtons(), 0, 13);
            Controls.Add(layout);
            FormClosing += SettingsWindowFormClosing;

            AcceptButton = _save;

            var cancel = GetCancelButton(layout);
            CancelButton = cancel;

            foreach (var preset in ModelSelectionPolicy.Presets)
            {
                _model.Items.Add(preset);
            }

            _endpoint.Text = current?.BaseUrl ?? string.Empty;
            _model.Text = string.IsNullOrWhiteSpace(current?.Model)
                ? ModelSelectionPolicy.RecommendedModel
                : current.Model;
            _apiKey.Text = current?.ApiKey ?? string.Empty;
            _allowInsecureHttp.Checked =
                current?.AllowInsecureHttp ?? false;
            UpdateModelGuidance();
            UpdateTransportWarning();
        }

        public AppSettings SavedSettings { get; private set; }

        protected override void OnFormClosed(
            FormClosedEventArgs eventArgs)
        {
            _checkCancellation?.Cancel();
            _checkCancellation?.Dispose();
            _checkCancellation = null;
            _client.Dispose();
            base.OnFormClosed(eventArgs);
        }

        private static Color SecondaryText
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

        private void ConfigureModelField()
        {
            _model.Dock = DockStyle.Fill;
            _model.DropDownStyle = ComboBoxStyle.DropDown;
            _model.FlatStyle = FlatStyle.Flat;
            _model.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _model.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _model.AutoCompleteSource = AutoCompleteSource.ListItems;
            _model.MaxDropDownItems = 12;
            _model.AccessibleName = "AI model";
            _model.AccessibleDescription =
                "Editable model identifier sent to the endpoint.";
            _model.TextChanged +=
                (sender, args) => UpdateModelGuidance();
        }

        private Control BuildModelRow()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 142));

            _useRecommended.Margin =
                new Padding(8, 0, 0, 0);
            _useRecommended.Click +=
                (sender, args) =>
                {
                    _model.Text =
                        ModelSelectionPolicy.RecommendedModel;
                    _model.Focus();
                };

            panel.Controls.Add(_model, 0, 0);
            panel.Controls.Add(_useRecommended, 1, 0);
            return panel;
        }

        private Control BuildButtons()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50));
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 50));

            var left = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            _checkEndpoint.Click += CheckEndpointClick;
            left.Controls.Add(_checkEndpoint);

            var right = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            _save.Click += SaveClick;
            var cancel = MakeButton("Cancel", false, 96);
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Name = "CancelSettings";
            right.Controls.Add(_save);
            right.Controls.Add(cancel);

            panel.Controls.Add(left, 0, 0);
            panel.Controls.Add(right, 1, 0);
            return panel;
        }

        private static Button GetCancelButton(
            Control root)
        {
            var matches = root.Controls.Find(
                "CancelSettings",
                true);
            return matches.Length > 0
                ? (Button)matches[0]
                : null;
        }

        private static void ConfigureField(
            TextBox field,
            string accessibleName,
            string accessibleDescription)
        {
            field.Dock = DockStyle.Fill;
            field.BorderStyle = BorderStyle.FixedSingle;
            field.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            field.AccessibleName = accessibleName;
            field.AccessibleDescription = accessibleDescription;
        }

        private static void ConfigureSupportingLabel(
            Label label)
        {
            label.AutoSize = true;
            label.MaximumSize = new Size(560, 0);
            label.ForeColor = SecondaryText;
        }

        private static Label FieldLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Font = new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    SystemFonts.MessageBoxFont.Size,
                    FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                Padding = new Padding(0, 0, 0, 4),
                Text = text
            };
        }

        private static Button MakeButton(
            string text,
            bool primary,
            int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(8, 0, 0, 0),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = primary
                ? (SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(0, 95, 184))
                : SystemColors.Window;
            button.ForeColor = primary
                ? SystemColors.HighlightText
                : SystemColors.WindowText;
            button.FlatAppearance.BorderColor = primary
                ? button.BackColor
                : SystemColors.ControlDark;
            button.AccessibleName = text;
            return button;
        }

        private AppSettings ReadFormSettings()
        {
            return new AppSettings
            {
                BaseUrl = _endpoint.Text,
                Model = _model.Text,
                ApiKey = _apiKey.Text,
                AllowInsecureHttp =
                    _allowInsecureHttp.Checked
            };
        }

        private async void CheckEndpointClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_checking)
            {
                _checkCancellation?.Cancel();
                return;
            }

            _error.Text = string.Empty;
            var settings = ReadFormSettings();
            if (!settings.IsConfigured)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint, " +
                    "model, and API key first.";
                return;
            }

            SetChecking(true);
            _checkCancellation =
                new CancellationTokenSource(
                    TimeSpan.FromSeconds(45));
            try
            {
                IReadOnlyList<string> models = null;
                var discoveryNote = string.Empty;
                try
                {
                    _testStatus.Text =
                        "Checking authentication and available models...";
                    models = await _client.GetModelsAsync(
                        settings,
                        _checkCancellation.Token);
                    AddDiscoveredModels(models);
                }
                catch (AiEndpointException exception)
                {
                    discoveryNote =
                        " Model discovery was unavailable [" +
                        exception.Code + "], so the entered model " +
                        "was tested directly.";
                    _testStatus.Text =
                        "Model discovery is unavailable. Testing the " +
                        "entered model directly...";
                }

                _testStatus.Text =
                    "Testing OpenAI-compatible mailbox tool calls...";
                var probe = ChatRequestFactory.Create(
                    settings.Model,
                    null,
                    new List<ChatTurn>(),
                    "Configuration check only. Call search_mailbox with query " +
                    "\"configuration-check\", folder \"inbox\", days_back 1, " +
                    "and max_results 1 before answering.");
                var response = await _client.CompleteAsync(
                    settings,
                    probe,
                    _checkCancellation.Token);
                var validCall = response.tool_calls != null &&
                    response.tool_calls.Any(call =>
                        call?.function != null &&
                        MailboxToolCatalog.IsApproved(
                            call.function.name));
                if (!validCall)
                {
                    throw new AiEndpointException(
                        "MODEL_TOOL_CALL_UNSUPPORTED",
                        "The endpoint answered, but this model did not return a " +
                        "compatible mailbox tool call. Choose a tool-capable model.");
                }

                var recommendation =
                    ModelSelectionPolicy.ChooseRecommended(models);
                var note =
                    recommendation.Length > 0 &&
                    !recommendation.Equals(
                        settings.Model,
                        StringComparison.OrdinalIgnoreCase)
                        ? " Best available balance: " +
                          recommendation + "."
                        : string.Empty;
                _testStatus.ForeColor = SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(20, 112, 70);
                _testStatus.Text =
                    "Endpoint verified. Authentication, model, and mailbox " +
                    "tool calling all passed." +
                    discoveryNote + note;
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[ENDPOINT_CHECK_CANCELLED] The endpoint check was cancelled or " +
                    "did not finish within 45 seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "ENDPOINT_CHECK_FAILED");
            }
            finally
            {
                _checkCancellation?.Dispose();
                _checkCancellation = null;
                SetChecking(false);
            }
        }

        private void AddDiscoveredModels(
            IEnumerable<string> models)
        {
            foreach (var model in models ??
                Enumerable.Empty<string>())
            {
                if (!_model.Items.Cast<object>()
                    .Any(item => string.Equals(
                        item.ToString(),
                        model,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    _model.Items.Add(model);
                }
            }
        }

        private void SetChecking(bool checking)
        {
            _checking = checking;
            _checkEndpoint.Text =
                checking ? "Cancel check" : "Check endpoint";
            _save.Enabled = !checking;
            _endpoint.Enabled = !checking;
            _model.Enabled = !checking;
            _apiKey.Enabled = !checking;
            _allowInsecureHttp.Enabled = !checking;
            _useRecommended.Enabled = !checking;
            if (checking)
            {
                _testStatus.ForeColor = SecondaryText;
            }
        }

        private void SaveClick(
            object sender,
            EventArgs eventArgs)
        {
            try
            {
                _error.Text = string.Empty;
                var settings = ReadFormSettings();
                _store.Save(settings);
                SavedSettings = settings;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "SETTINGS_SAVE_FAILED");
            }
        }

        private void SettingsWindowFormClosing(
            object sender,
            FormClosingEventArgs eventArgs)
        {
            if (!_checking)
            {
                return;
            }

            eventArgs.Cancel = true;
            _checkCancellation?.Cancel();
            _testStatus.Text =
                "Cancelling the endpoint check...";
        }

        private void InsecureHttpChanged(
            object sender,
            EventArgs eventArgs)
        {
            UpdateTransportWarning();
        }

        private void UpdateModelGuidance()
        {
            _modelGuidance.Text =
                ModelSelectionPolicy.DescriptionFor(_model.Text);
        }

        private void UpdateTransportWarning()
        {
            if (_allowInsecureHttp.Checked)
            {
                _transportWarning.ForeColor = ErrorText;
                _transportWarning.Text =
                    "Warning: with HTTP, the API key, prompts, and retrieved email " +
                    "context cross the network without transport encryption.";
                return;
            }

            _transportWarning.ForeColor = SecondaryText;
            _transportWarning.Text =
                "Loopback HTTP remains available without this setting. " +
                "HTTPS is recommended for every remote endpoint.";
        }
    }
}

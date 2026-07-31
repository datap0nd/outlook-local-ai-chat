using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.UI
{
    public sealed class SettingsWindow : Form
    {
        private readonly TextBox _endpoint = new TextBox();
        private readonly ComboBox _model = new ComboBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly CheckBox _allowInsecureHttp = new CheckBox();
        private readonly Label _transportWarning = new Label();
        private readonly Label _modelGuidance = new Label();
        private readonly Label _testStatus = new Label();
        private readonly CheckBox _useToneProfile = new CheckBox();
        private readonly RichTextBox _toneProfile = new RichTextBox();
        private readonly Label _toneStatus = new Label();
        private readonly Label _error = new Label();
        private readonly Button _checkEndpoint =
            MakeButton("Check endpoint", false, 128);
        private readonly Button _useRecommended =
            MakeButton("Use recommended", false, 134);
        private readonly Button _analyzeTone =
            MakeButton("Analyze 15 sent emails", false, 176);
        private readonly Button _save =
            MakeButton("Save", true, 96);
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly SettingsStore _store;
        private readonly object _outlookApplication;
        private CancellationTokenSource _checkCancellation;
        private CancellationTokenSource _toneCancellation;
        private bool _checking;
        private bool _analyzingTone;

        public SettingsWindow(
            SettingsStore store,
            AppSettings current)
            : this(store, current, null)
        {
        }

        public SettingsWindow(
            SettingsStore store,
            AppSettings current,
            object outlookApplication)
        {
            _store = store ??
                throw new ArgumentNullException(nameof(store));
            _outlookApplication = outlookApplication;

            Text = "MailAI settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(700, 670);
            MinimumSize = new Size(620, 620);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;

            ConfigureFields();
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18, 16, 18, 14)
            };
            root.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 46));

            var tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                AccessibleName = "MailAI settings sections"
            };
            tabs.TabPages.Add(BuildConnectionPage());
            tabs.TabPages.Add(BuildWritingStylePage());
            root.Controls.Add(tabs, 0, 0);

            ConfigureSupportingLabel(_error);
            _error.ForeColor = ErrorText;
            _error.AccessibleName = "Settings error";
            _error.AccessibleRole = AccessibleRole.Alert;
            root.Controls.Add(_error, 0, 1);
            var buttons = BuildButtons();
            root.Controls.Add(buttons, 0, 2);
            Controls.Add(root);
            FormClosing += SettingsWindowFormClosing;

            AcceptButton = _save;
            CancelButton = GetCancelButton(buttons);

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
            _toneProfile.Text = TextBoundary.PlainText(
                current?.ToneProfile,
                TextBoundary.MaxToneProfileCharacters);
            _useToneProfile.Checked =
                (current?.UseToneProfile ?? false) &&
                _toneProfile.TextLength > 0;
            UpdateModelGuidance();
            UpdateTransportWarning();
        }

        public AppSettings SavedSettings { get; private set; }

        protected override void OnFormClosed(
            FormClosedEventArgs eventArgs)
        {
            _checkCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _checkCancellation?.Dispose();
            _toneCancellation?.Dispose();
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

        private static Color SuccessText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(20, 112, 70);
            }
        }

        private void ConfigureFields()
        {
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

            _allowInsecureHttp.AutoSize = true;
            _allowInsecureHttp.Text =
                "Allow insecure HTTP for non-local endpoints";
            _allowInsecureHttp.AccessibleName =
                "Allow insecure HTTP";
            _allowInsecureHttp.AccessibleDescription =
                "Allows the API key, prompts, and email context to be sent " +
                "without transport encryption.";
            _allowInsecureHttp.CheckedChanged += InsecureHttpChanged;

            _useToneProfile.AutoSize = true;
            _useToneProfile.Text =
                "Use this writing profile for drafts";
            _useToneProfile.AccessibleDescription =
                "Applies the editable writing profile only when creating or updating drafts.";

            _toneProfile.Dock = DockStyle.Fill;
            _toneProfile.BorderStyle = BorderStyle.FixedSingle;
            _toneProfile.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _toneProfile.MaxLength =
                TextBoundary.MaxToneProfileCharacters;
            _toneProfile.ScrollBars =
                RichTextBoxScrollBars.Vertical;
            _toneProfile.DetectUrls = false;
            _toneProfile.AccessibleName =
                "Editable email writing profile";
        }

        private TabPage BuildConnectionPage()
        {
            var page = new TabPage("Connection");
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 12,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            layout.Controls.Add(FieldLabel("Endpoint or base URL"), 0, 0);
            layout.Controls.Add(_endpoint, 0, 1);
            layout.Controls.Add(FieldLabel("Model"), 0, 2);
            layout.Controls.Add(BuildModelRow(), 0, 3);
            ConfigureSupportingLabel(_modelGuidance);
            layout.Controls.Add(_modelGuidance, 0, 4);
            layout.Controls.Add(FieldLabel("API key"), 0, 5);
            layout.Controls.Add(_apiKey, 0, 6);
            layout.Controls.Add(_allowInsecureHttp, 0, 7);
            ConfigureSupportingLabel(_transportWarning);
            _transportWarning.AccessibleRole = AccessibleRole.Alert;
            layout.Controls.Add(_transportWarning, 0, 8);

            var hint = SupportingText(
                "The endpoint receives your prompt, conversation, and only bounded " +
                "context. MailAI exposes read tools and guarded draft creation only. " +
                "It has no send, move, delete, or mailbox mutation capability.");
            layout.Controls.Add(hint, 0, 9);
            ConfigureSupportingLabel(_testStatus);
            _testStatus.Text =
                "Check the endpoint to verify authentication, model discovery, and tool-call compatibility.";
            _testStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_testStatus, 0, 10);

            _checkEndpoint.Click += CheckEndpointClick;
            var checkRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0)
            };
            checkRow.Controls.Add(_checkEndpoint);
            layout.Controls.Add(checkRow, 0, 11);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildWritingStylePage()
        {
            var page = new TabPage("Writing style");
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 8,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var disclosure = SupportingText(
                "Nothing is analyzed automatically. When you click Analyze, MailAI " +
                "reads up to 15 recent Sent Items messages, removes obvious quoted " +
                "history, and sends bounded samples to your configured endpoint. " +
                "Review and edit the result before saving.");
            disclosure.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(disclosure, 0, 0);
            layout.Controls.Add(_useToneProfile, 0, 1);
            layout.Controls.Add(FieldLabel("My drafting instructions"), 0, 2);
            layout.Controls.Add(_toneProfile, 0, 3);

            var scope = SupportingText(
                "The profile affects wording, greeting, cadence, and sign-off only. " +
                "It cannot change MailAI permissions or security rules.");
            layout.Controls.Add(scope, 0, 4);

            _analyzeTone.Click += AnalyzeToneClick;
            var actionRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0)
            };
            actionRow.Controls.Add(_analyzeTone);
            layout.Controls.Add(actionRow, 0, 5);

            ConfigureSupportingLabel(_toneStatus);
            _toneStatus.Text =
                "Analysis requires at least five usable sent messages and never runs without this button.";
            _toneStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_toneStatus, 0, 6);
            layout.Controls.Add(
                SupportingText(
                    "Tip: keep the profile general. Remove names, client details, " +
                    "project facts, and anything you would not want reused."),
                0,
                7);
            page.Controls.Add(layout);
            return page;
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
            _useRecommended.Margin = new Padding(8, 0, 0, 0);
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
            var panel = new FlowLayoutPanel
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
            panel.Controls.Add(_save);
            panel.Controls.Add(cancel);
            return panel;
        }

        private static Button GetCancelButton(Control root)
        {
            var matches = root.Controls.Find("CancelSettings", true);
            return matches.Length > 0 ? (Button)matches[0] : null;
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

        private static void ConfigureSupportingLabel(Label label)
        {
            label.AutoSize = true;
            label.MaximumSize = new Size(620, 0);
            label.ForeColor = SecondaryText;
        }

        private static Label SupportingText(string text)
        {
            var label = new Label { Text = text };
            ConfigureSupportingLabel(label);
            return label;
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
            var profile = TextBoundary.PlainText(
                _toneProfile.Text,
                TextBoundary.MaxToneProfileCharacters);
            return new AppSettings
            {
                BaseUrl = _endpoint.Text,
                Model = _model.Text,
                ApiKey = _apiKey.Text,
                AllowInsecureHttp = _allowInsecureHttp.Checked,
                ToneProfile = profile,
                UseToneProfile =
                    _useToneProfile.Checked && profile.Length > 0
            };
        }

        private async void AnalyzeToneClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_analyzingTone)
            {
                _toneCancellation?.Cancel();
                return;
            }

            _error.Text = string.Empty;
            if (_outlookApplication == null)
            {
                _error.Text =
                    "[OUTLOOK_NOT_READY] Open MailAI from Outlook before analyzing your writing style.";
                return;
            }

            var settings = ReadFormSettings();
            if (!settings.IsConfigured)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Configure the endpoint, model, and API key first.";
                return;
            }

            SetToneAnalyzing(true);
            _toneCancellation = new CancellationTokenSource(
                TimeSpan.FromSeconds(120));
            try
            {
                _toneStatus.Text =
                    "Reading and cleaning up to 15 recent Sent Items messages...";
                var samples = new SentMailToneSampler(
                    _outlookApplication).CaptureRecent();
                if (samples.Count < 5)
                {
                    throw new AiEndpointException(
                        "TONE_SAMPLES_INSUFFICIENT",
                        "At least five usable sent messages are required. Found " +
                        samples.Count + ".");
                }

                _toneStatus.Text =
                    "Analyzing " + samples.Count +
                    " bounded sent-email samples with " + settings.Model + "...";
                var request = ToneProfileRequestFactory.Create(
                    settings.Model,
                    samples);
                var response = await _client.CompleteAsync(
                    settings,
                    request,
                    _toneCancellation.Token);
                if (response.tool_calls != null &&
                    response.tool_calls.Count > 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_INVALID",
                        "The model returned tool calls during style analysis.");
                }

                var profile = SafeModelText.Format(
                    response.content,
                    TextBoundary.MaxToneProfileCharacters).PlainText;
                if (profile.Length == 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_EMPTY",
                        "The model returned an empty writing profile.");
                }

                _toneProfile.Text = profile;
                _useToneProfile.Checked = true;
                _toneStatus.ForeColor = SuccessText;
                _toneStatus.Text =
                    "Writing profile generated from " + samples.Count +
                    " sent messages. Review and edit it, then click Save.";
            }
            catch (OperationCanceledException)
            {
                _toneStatus.Text =
                    "Writing-style analysis was cancelled or timed out after 120 seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "TONE_ANALYSIS_FAILED");
            }
            finally
            {
                _toneCancellation?.Dispose();
                _toneCancellation = null;
                SetToneAnalyzing(false);
            }
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
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint, model, and API key first.";
                return;
            }

            SetChecking(true);
            _checkCancellation = new CancellationTokenSource(
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
                        " Model discovery was unavailable [" + exception.Code +
                        "], so the entered model was tested directly.";
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
                        "The endpoint answered, but this model did not return a compatible mailbox tool call.");
                }

                var recommendation =
                    ModelSelectionPolicy.ChooseRecommended(models);
                var note = recommendation.Length > 0 &&
                    !recommendation.Equals(
                        settings.Model,
                        StringComparison.OrdinalIgnoreCase)
                        ? " Best available balance: " + recommendation + "."
                        : string.Empty;
                _testStatus.ForeColor = SuccessText;
                _testStatus.Text =
                    "Endpoint verified. Authentication, model, and mailbox tool calling passed." +
                    discoveryNote + note;
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[ENDPOINT_CHECK_CANCELLED] The endpoint check was cancelled or timed out after 45 seconds.";
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

        private void AddDiscoveredModels(IEnumerable<string> models)
        {
            foreach (var model in models ?? Enumerable.Empty<string>())
            {
                if (!_model.Items.Cast<object>().Any(item =>
                    string.Equals(
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
            SetCommonControlsEnabled(!checking && !_analyzingTone);
            if (checking)
            {
                _testStatus.ForeColor = SecondaryText;
            }
        }

        private void SetToneAnalyzing(bool analyzing)
        {
            _analyzingTone = analyzing;
            _analyzeTone.Text =
                analyzing ? "Cancel analysis" : "Analyze 15 sent emails";
            SetCommonControlsEnabled(!analyzing && !_checking);
            _analyzeTone.Enabled = analyzing || !_checking;
            if (analyzing)
            {
                _toneStatus.ForeColor = SecondaryText;
            }
        }

        private void SetCommonControlsEnabled(bool enabled)
        {
            _save.Enabled = enabled;
            _endpoint.Enabled = enabled;
            _model.Enabled = enabled;
            _apiKey.Enabled = enabled;
            _allowInsecureHttp.Enabled = enabled;
            _useRecommended.Enabled = enabled;
            _useToneProfile.Enabled = enabled;
            _toneProfile.Enabled = enabled;
            _checkEndpoint.Enabled = enabled || _checking;
            _analyzeTone.Enabled = enabled || _analyzingTone;
        }

        private void SaveClick(object sender, EventArgs eventArgs)
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
            if (!_checking && !_analyzingTone)
            {
                return;
            }

            eventArgs.Cancel = true;
            _checkCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _error.Text =
                "Cancelling the active settings operation. Close again when it finishes.";
        }

        private void InsecureHttpChanged(object sender, EventArgs eventArgs)
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

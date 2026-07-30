using System;
using System.Drawing;
using System.Windows.Forms;
using OutlookLocalAIChat.Configuration;

namespace OutlookLocalAIChat.UI
{
    public sealed class SettingsWindow : Form
    {
        private readonly TextBox _endpoint = new TextBox();
        private readonly TextBox _model = new TextBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly Label _error = new Label();
        private readonly SettingsStore _store;

        public SettingsWindow(SettingsStore store, AppSettings current)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));

            Text = "AI endpoint settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 365);
            MinimumSize = new Size(480, 390);
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
                RowCount = 10,
                Padding = new Padding(24, 20, 24, 20)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));

            ConfigureField(
                _endpoint,
                "AI endpoint",
                "HTTPS endpoint or loopback HTTP base URL.");
            ConfigureField(_model, "Model name", "Model identifier sent to the endpoint.");
            ConfigureField(_apiKey, "API key", "Encrypted for the current Windows user.");
            _apiKey.UseSystemPasswordChar = true;

            layout.Controls.Add(FieldLabel("Endpoint or base URL"), 0, 0);
            layout.Controls.Add(_endpoint, 0, 1);
            layout.Controls.Add(FieldLabel("Model"), 0, 2);
            layout.Controls.Add(_model, 0, 3);
            layout.Controls.Add(FieldLabel("API key"), 0, 4);
            layout.Controls.Add(_apiKey, 0, 5);

            var hint = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(460, 0),
                ForeColor = SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(80, 80, 80),
                Text =
                    "Your prompt and recent conversation are sent to this endpoint. " +
                    "The model can request bounded read-only context from your local " +
                    "Inbox and Sent Items. Use HTTPS, or HTTP only for a local endpoint. " +
                    "The key is encrypted for your Windows account.",
                AccessibleName = "Endpoint data disclosure"
            };
            layout.Controls.Add(hint, 0, 6);

            _error.AutoSize = true;
            _error.ForeColor = SystemInformation.HighContrast
                ? SystemColors.HotTrack
                : Color.FromArgb(163, 38, 38);
            _error.Padding = new Padding(0, 10, 0, 0);
            _error.AccessibleName = "Settings error";
            _error.AccessibleRole = AccessibleRole.Alert;
            layout.Controls.Add(_error, 0, 7);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 5, 0, 0)
            };
            var save = MakeButton("Save", true);
            save.Click += SaveClick;
            var cancel = MakeButton("Cancel", false);
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);
            layout.Controls.Add(buttons, 0, 9);

            Controls.Add(layout);
            AcceptButton = save;
            CancelButton = cancel;

            _endpoint.Text = current?.BaseUrl ?? string.Empty;
            _model.Text = current?.Model ?? string.Empty;
            _apiKey.Text = current?.ApiKey ?? string.Empty;
        }

        public AppSettings SavedSettings { get; private set; }

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

        private static Button MakeButton(string text, bool primary)
        {
            var button = new Button
            {
                Text = text,
                Width = 96,
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

        private void SaveClick(object sender, EventArgs eventArgs)
        {
            try
            {
                var settings = new AppSettings
                {
                    BaseUrl = _endpoint.Text,
                    Model = _model.Text,
                    ApiKey = _apiKey.Text
                };
                _store.Save(settings);
                SavedSettings = settings;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                _error.Text = exception.Message;
            }
        }
    }
}

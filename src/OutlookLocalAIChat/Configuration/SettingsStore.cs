using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookLocalAIChat.Configuration
{
    public sealed class SettingsStore
    {
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("OutlookLocalAIChat.Settings.v1");

        private readonly string _settingsPath;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public SettingsStore()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "OutlookLocalAIChat");
            _settingsPath = Path.Combine(directory, "settings.json");
        }

        public AppSettings Load()
        {
            try
            {
                if (!File.Exists(_settingsPath))
                {
                    return new AppSettings();
                }

                var stored = _serializer.Deserialize<StoredSettings>(
                    File.ReadAllText(_settingsPath, Encoding.UTF8));
                if (stored == null)
                {
                    return new AppSettings();
                }

                return new AppSettings
                {
                    BaseUrl = stored.BaseUrl ?? string.Empty,
                    Model = string.IsNullOrWhiteSpace(stored.Model)
                        ? ModelSelectionPolicy.RecommendedModel
                        : stored.Model,
                    ApiKey = Unprotect(stored.ProtectedApiKey),
                    AllowInsecureHttp = stored.AllowInsecureHttp
                };
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            Uri endpoint;
            if (!AppSettings.TryGetChatCompletionsUri(
                settings.BaseUrl,
                settings.AllowInsecureHttp,
                out endpoint))
            {
                throw new InvalidOperationException(
                    "Use HTTPS, loopback HTTP, or explicitly allow insecure HTTP.");
            }

            if (settings.Model.Trim().Length == 0)
            {
                throw new InvalidOperationException("Enter a model name.");
            }

            if (settings.ApiKey.Trim().Length == 0)
            {
                throw new InvalidOperationException("Enter an API key.");
            }

            var directory = Path.GetDirectoryName(_settingsPath);
            Directory.CreateDirectory(directory);

            var stored = new StoredSettings
            {
                BaseUrl = settings.BaseUrl.Trim(),
                Model = settings.Model.Trim(),
                ProtectedApiKey = Protect(settings.ApiKey.Trim()),
                AllowInsecureHttp = settings.AllowInsecureHttp
            };

            File.WriteAllText(
                _settingsPath,
                _serializer.Serialize(stored),
                new UTF8Encoding(false));
        }

        private static string Protect(string value)
        {
            var clearBytes = Encoding.UTF8.GetBytes(value);
            var protectedBytes = ProtectedData.Protect(
                clearBytes,
                Entropy,
                DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        private static string Unprotect(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var protectedBytes = Convert.FromBase64String(value);
            var clearBytes = ProtectedData.Unprotect(
                protectedBytes,
                Entropy,
                DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(clearBytes);
        }

        private sealed class StoredSettings
        {
            public string BaseUrl { get; set; }

            public string Model { get; set; }

            public string ProtectedApiKey { get; set; }

            public bool AllowInsecureHttp { get; set; }
        }
    }
}

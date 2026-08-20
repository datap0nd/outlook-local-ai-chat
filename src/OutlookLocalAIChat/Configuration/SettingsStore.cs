using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Security;

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
                    Model = stored.Model ?? string.Empty,
                    ApiKey = Unprotect(stored.ProtectedApiKey),
                    AllowInsecureHttp = stored.AllowInsecureHttp,
                    UseGeminiSignIn = stored.UseGeminiSignIn,
                    GeminiRefreshToken = Unprotect(
                        stored.ProtectedGeminiRefreshToken),
                    GeminiProject = TextBoundary.SingleLine(
                        stored.GeminiProject,
                        200),
                    ToneProfile = TextBoundary.PlainText(
                        stored.ToneProfile,
                        TextBoundary.MaxToneProfileCharacters),
                    UseToneProfile = stored.UseToneProfile,
                    ToneStrength = Math.Max(
                        10,
                        Math.Min(
                            100,
                            stored.ToneStrength == 0
                                ? 60
                                : stored.ToneStrength)),
                    DraftRules = TextBoundary.PlainText(
                        stored.DraftRules,
                        2000),
                    SwitchToVisionModelForImages =
                        stored.SwitchToVisionModelForImages,
                    DiscoveredModels = NormalizeDiscoveredModels(
                        stored.DiscoveredModels),
                    // A missing UseCustomLimits (older settings
                    // files) means recommended limits; missing
                    // custom values (0) fall back to recommended.
                    UseRecommendedLimits = !stored.UseCustomLimits,
                    LimitContextMultiplier = OrDefault(
                        stored.LimitContextMultiplier,
                        1),
                    LimitPromptCharacters = OrDefault(
                        stored.LimitPromptCharacters,
                        TextBoundary
                            .RecommendedUserPromptCharacters),
                    LimitAssistantCharacters = OrDefault(
                        stored.LimitAssistantCharacters,
                        TextBoundary
                            .RecommendedAssistantCharacters),
                    LimitHistoryTurns = OrDefault(
                        stored.LimitHistoryTurns,
                        TextBoundary
                            .RecommendedConversationTurns),
                    LimitToolRounds = OrDefault(
                        stored.LimitToolRounds,
                        TextBoundary.RecommendedToolRounds),
                    LimitToolCallsPerRound = OrDefault(
                        stored.LimitToolCallsPerRound,
                        TextBoundary
                            .RecommendedToolCallsPerRound)
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

            if (settings.Model.Trim().Length == 0)
            {
                throw new InvalidOperationException("Enter a model name.");
            }

            if (!settings.UseGeminiSignIn)
            {
                Uri endpoint;
                if (!AppSettings.TryGetChatCompletionsUri(
                    settings.BaseUrl,
                    settings.AllowInsecureHttp,
                    out endpoint))
                {
                    throw new InvalidOperationException(
                        "Use HTTPS, loopback HTTP, or explicitly allow insecure HTTP.");
                }

                if (settings.ApiKey.Trim().Length == 0)
                {
                    throw new InvalidOperationException("Enter an API key.");
                }
            }

            var directory = Path.GetDirectoryName(_settingsPath);
            Directory.CreateDirectory(directory);

            var stored = new StoredSettings
            {
                BaseUrl = settings.BaseUrl.Trim(),
                Model = settings.Model.Trim(),
                ProtectedApiKey =
                    settings.ApiKey.Trim().Length > 0
                        ? Protect(settings.ApiKey.Trim())
                        : string.Empty,
                AllowInsecureHttp = settings.AllowInsecureHttp,
                UseGeminiSignIn = settings.UseGeminiSignIn,
                ProtectedGeminiRefreshToken =
                    settings.GeminiRefreshToken.Trim().Length > 0
                        ? Protect(
                            settings.GeminiRefreshToken.Trim())
                        : string.Empty,
                GeminiProject = TextBoundary.SingleLine(
                    settings.GeminiProject,
                    200),
                ToneProfile = TextBoundary.PlainText(
                    settings.ToneProfile,
                    TextBoundary.MaxToneProfileCharacters),
                UseToneProfile = settings.UseToneProfile &&
                    !string.IsNullOrWhiteSpace(settings.ToneProfile),
                ToneStrength = Math.Max(
                    10,
                    Math.Min(100, settings.ToneStrength)),
                DraftRules = TextBoundary.PlainText(
                    settings.DraftRules,
                    2000),
                SwitchToVisionModelForImages =
                    settings.SwitchToVisionModelForImages,
                DiscoveredModels = NormalizeDiscoveredModels(
                    settings.DiscoveredModels),
                UseCustomLimits = !settings.UseRecommendedLimits,
                LimitContextMultiplier =
                    settings.LimitContextMultiplier,
                LimitPromptCharacters =
                    settings.LimitPromptCharacters,
                LimitAssistantCharacters =
                    settings.LimitAssistantCharacters,
                LimitHistoryTurns = settings.LimitHistoryTurns,
                LimitToolRounds = settings.LimitToolRounds,
                LimitToolCallsPerRound =
                    settings.LimitToolCallsPerRound
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

        private static List<string> NormalizeDiscoveredModels(
            IEnumerable<string> models)
        {
            return (models ?? Enumerable.Empty<string>())
                .Select(model => TextBoundary.PlainText(model, 200))
                .Where(model => model.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(model => model, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private sealed class StoredSettings
        {
            public string BaseUrl { get; set; }

            public string Model { get; set; }

            public string ProtectedApiKey { get; set; }

            public bool AllowInsecureHttp { get; set; }

            public bool UseGeminiSignIn { get; set; }

            public string ProtectedGeminiRefreshToken { get; set; }

            public string GeminiProject { get; set; }

            public string ToneProfile { get; set; }

            public bool UseToneProfile { get; set; }

            public int ToneStrength { get; set; }

            public string DraftRules { get; set; }

            public bool SwitchToVisionModelForImages { get; set; }

            public List<string> DiscoveredModels { get; set; }

            public bool UseCustomLimits { get; set; }

            public int LimitContextMultiplier { get; set; }

            public int LimitPromptCharacters { get; set; }

            public int LimitAssistantCharacters { get; set; }

            public int LimitHistoryTurns { get; set; }

            public int LimitToolRounds { get; set; }

            public int LimitToolCallsPerRound { get; set; }
        }

        private static int OrDefault(int value, int fallback)
        {
            return value > 0 ? value : fallback;
        }
    }
}

using System;

namespace OutlookLocalAIChat.Configuration
{
    public sealed class AppSettings
    {
        public string BaseUrl { get; set; } = string.Empty;

        public string Model { get; set; } = string.Empty;

        public string ApiKey { get; set; } = string.Empty;

        public bool AllowInsecureHttp { get; set; }

        public bool IsConfigured
        {
            get
            {
                Uri endpoint;
                return Model.Trim().Length > 0 &&
                       ApiKey.Trim().Length > 0 &&
                       TryGetChatCompletionsUri(
                           BaseUrl,
                           AllowInsecureHttp,
                           out endpoint);
            }
        }

        public static bool TryGetChatCompletionsUri(string value, out Uri endpoint)
        {
            return TryGetChatCompletionsUri(
                value,
                false,
                out endpoint);
        }

        public static bool TryGetChatCompletionsUri(
            string value,
            bool allowInsecureHttp,
            out Uri endpoint)
        {
            endpoint = null;
            Uri baseUri;
            if (!Uri.TryCreate(value?.Trim(), UriKind.Absolute, out baseUri))
            {
                return false;
            }

            var isHttps = baseUri.Scheme.Equals(
                Uri.UriSchemeHttps,
                StringComparison.OrdinalIgnoreCase);
            var isAllowedHttp =
                baseUri.Scheme.Equals(
                    Uri.UriSchemeHttp,
                    StringComparison.OrdinalIgnoreCase) &&
                (baseUri.IsLoopback || allowInsecureHttp);

            if (!isHttps && !isAllowedHttp)
            {
                return false;
            }

            var path = baseUri.AbsolutePath.TrimEnd('/');
            string completedPath;
            if (path.EndsWith(
                "/chat/completions",
                StringComparison.OrdinalIgnoreCase))
            {
                completedPath = path;
            }
            else if (path.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                completedPath = path + "/chat/completions";
            }
            else
            {
                completedPath = path + "/v1/chat/completions";
            }

            var builder = new UriBuilder(baseUri)
            {
                Path = completedPath,
                Query = string.Empty,
                Fragment = string.Empty
            };

            endpoint = builder.Uri;
            return true;
        }
    }
}

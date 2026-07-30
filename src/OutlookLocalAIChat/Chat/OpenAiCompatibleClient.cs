using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Chat
{
    public sealed class OpenAiCompatibleClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public OpenAiCompatibleClient()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
        }

        public async Task<ChatCompletionResponseMessage> CompleteAsync(
            AppSettings settings,
            ChatCompletionRequest requestModel,
            CancellationToken cancellationToken)
        {
            if (settings == null || !settings.IsConfigured)
            {
                throw new AiEndpointException(
                    "CONFIGURATION_INCOMPLETE",
                    "Open Settings and configure the endpoint, model, and API key.");
            }

            if (requestModel == null)
            {
                throw new ArgumentNullException(nameof(requestModel));
            }

            Uri endpoint;
            if (!AppSettings.TryGetChatCompletionsUri(
                settings.BaseUrl,
                out endpoint))
            {
                throw new AiEndpointException(
                    "ENDPOINT_INVALID",
                    "The configured endpoint is invalid.");
            }

            var requestJson = _serializer.Serialize(requestModel);

            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                request.Content = new StringContent(
                    requestJson,
                    Encoding.UTF8,
                    "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await _httpClient
                        .SendAsync(
                            request,
                            HttpCompletionOption.ResponseHeadersRead,
                            cancellationToken)
                        .ConfigureAwait(true);
                }
                catch (OperationCanceledException exception)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }

                    throw new AiEndpointException(
                        "AI_TIMEOUT",
                        "The AI endpoint did not respond within 120 seconds.",
                        exception);
                }
                catch (HttpRequestException exception)
                {
                    throw CreateNetworkException(
                        exception,
                        endpoint);
                }

                using (response)
                {
                    string responseText;
                    try
                    {
                        responseText = await ReadBoundedAsync(
                            response.Content,
                            cancellationToken).ConfigureAwait(true);
                    }
                    catch (AiEndpointException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_READ_FAILED",
                            "The AI endpoint response could not be read.",
                            exception);
                    }

                    var requestId = GetRequestId(response);
                    if (!response.IsSuccessStatusCode)
                    {
                        var error = TryReadError(responseText);
                        var status = (int)response.StatusCode;
                        var reason = string.IsNullOrWhiteSpace(response.ReasonPhrase)
                            ? response.StatusCode.ToString()
                            : response.ReasonPhrase;
                        throw new AiEndpointException(
                            BuildHttpCode(status, reason),
                            "The AI endpoint rejected the request: " +
                            status + " " + reason + "." +
                            RecoveryHint(status),
                            httpStatus: status,
                            providerCode: error?.code ?? error?.type,
                            requestId: requestId,
                            responseSnippet: error?.message ?? responseText);
                    }

                    ChatCompletionResponse completion;
                    try
                    {
                        completion =
                            _serializer.Deserialize<ChatCompletionResponse>(
                                responseText);
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_INVALID_JSON",
                            "The AI endpoint returned invalid JSON.",
                            exception,
                            (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    var message =
                        completion?.choices != null &&
                        completion.choices.Count > 0
                            ? completion.choices[0]?.message
                            : null;

                    var hasToolCalls =
                        message?.tool_calls != null &&
                        message.tool_calls.Count > 0;
                    if (message == null ||
                        (!hasToolCalls &&
                         string.IsNullOrWhiteSpace(message.content)))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The AI endpoint returned neither message text nor tool calls.",
                            httpStatus: (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    message.content = TextBoundary.PlainText(
                        message.content,
                        TextBoundary.MaxAssistantCharacters);
                    return message;
                }
            }
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }

        private static async Task<string> ReadBoundedAsync(
            HttpContent content,
            CancellationToken cancellationToken)
        {
            if (content.Headers.ContentLength.HasValue &&
                content.Headers.ContentLength.Value >
                TextBoundary.MaxHttpResponseCharacters)
            {
                throw new AiEndpointException(
                    "RESPONSE_TOO_LARGE",
                    "The AI endpoint response was too large.");
            }

            using (var stream = await content.ReadAsStreamAsync()
                .ConfigureAwait(true))
            using (var reader = new StreamReader(
                stream,
                Encoding.UTF8,
                true,
                4096,
                false))
            {
                var builder = new StringBuilder();
                var buffer = new char[4096];
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var read = await reader.ReadAsync(
                        buffer,
                        0,
                        buffer.Length).ConfigureAwait(true);
                    if (read == 0)
                    {
                        break;
                    }

                    if (builder.Length + read >
                        TextBoundary.MaxHttpResponseCharacters)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_TOO_LARGE",
                            "The AI endpoint response was too large.");
                    }

                    builder.Append(buffer, 0, read);
                }

                return builder.ToString();
            }
        }

        private static AiEndpointException CreateNetworkException(
            HttpRequestException exception,
            Uri endpoint)
        {
            var webException = FindWebException(exception);
            var code = "NETWORK_REQUEST_FAILED";
            if (webException != null)
            {
                switch (webException.Status)
                {
                    case WebExceptionStatus.NameResolutionFailure:
                        code = "NETWORK_NAME_RESOLUTION";
                        break;
                    case WebExceptionStatus.ConnectFailure:
                        code = "NETWORK_CONNECT_FAILURE";
                        break;
                    case WebExceptionStatus.TrustFailure:
                    case WebExceptionStatus.SecureChannelFailure:
                        code = "TLS_SECURE_CHANNEL_FAILURE";
                        break;
                    case WebExceptionStatus.ProxyNameResolutionFailure:
                        code = "NETWORK_PROXY_NAME_RESOLUTION";
                        break;
                    case WebExceptionStatus.Timeout:
                        code = "AI_TIMEOUT";
                        break;
                    default:
                        code = "NETWORK_" +
                            webException.Status.ToString().ToUpperInvariant();
                        break;
                }
            }

            return new AiEndpointException(
                code,
                "The AI endpoint could not be reached. Check its URL, " +
                "network access, TLS certificate, and whether the local server is running.",
                exception,
                transportDetails: BuildTransportDetails(
                    exception,
                    endpoint));
        }

        private static string BuildTransportDetails(
            Exception exception,
            Uri endpoint)
        {
            var details = new List<string>();
            if (endpoint != null)
            {
                details.Add(
                    "Target " +
                    endpoint.Scheme +
                    "://" +
                    endpoint.Host +
                    ":" +
                    endpoint.Port);
            }

            var current = exception;
            var depth = 0;
            while (current != null && depth < 8)
            {
                var item =
                    current.GetType().Name +
                    " HRESULT 0x" +
                    current.HResult.ToString("X8");

                var webException = current as WebException;
                if (webException != null)
                {
                    item +=
                        " WebExceptionStatus " +
                        webException.Status;
                }

                var socketException = current as SocketException;
                if (socketException != null)
                {
                    item +=
                        " SocketError " +
                        socketException.SocketErrorCode +
                        " NativeError " +
                        socketException.NativeErrorCode;
                }

                var message = TextBoundary.PlainText(
                    current.Message,
                    600)
                    .Replace("\r", " ")
                    .Replace("\n", " ")
                    .Trim();
                if (message.Length > 0)
                {
                    item += ": " + message;
                }

                details.Add(item);
                current = current.InnerException;
                depth++;
            }

            return string.Join(" | ", details);
        }

        private ChatCompletionError TryReadError(string responseText)
        {
            try
            {
                return _serializer
                    .Deserialize<ChatCompletionResponse>(responseText)
                    ?.error;
            }
            catch
            {
                return null;
            }
        }

        private static WebException FindWebException(Exception exception)
        {
            var current = exception;
            while (current != null)
            {
                var webException = current as WebException;
                if (webException != null)
                {
                    return webException;
                }

                current = current.InnerException;
            }

            return null;
        }

        private static string GetRequestId(HttpResponseMessage response)
        {
            foreach (var name in new[]
            {
                "x-request-id",
                "request-id",
                "x-correlation-id",
                "traceparent"
            })
            {
                IEnumerable<string> values;
                if (response.Headers.TryGetValues(name, out values))
                {
                    foreach (var value in values)
                    {
                        return TextBoundary.PlainText(value, 200);
                    }
                }
            }

            return string.Empty;
        }

        private static string BuildHttpCode(int status, string reason)
        {
            var builder = new StringBuilder();
            foreach (var character in reason ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character))
                {
                    builder.Append(char.ToUpperInvariant(character));
                }
                else if (builder.Length > 0 &&
                         builder[builder.Length - 1] != '_')
                {
                    builder.Append('_');
                }
            }

            return
                "HTTP_" +
                status +
                "_" +
                builder.ToString().Trim('_');
        }

        private static string RecoveryHint(int status)
        {
            switch (status)
            {
                case 400:
                    return
                        " The endpoint may not support OpenAI-compatible tool calling, " +
                        "or the model name/request format is invalid.";
                case 401:
                    return " Verify the API key.";
                case 403:
                    return " Verify the API key permissions and endpoint policy.";
                case 404:
                    return " Verify the base URL and model name.";
                case 408:
                    return " Retry after checking endpoint load.";
                case 413:
                    return " The endpoint rejected the bounded mailbox context size.";
                case 429:
                    return " The endpoint is rate-limiting requests. Retry later.";
                default:
                    return status >= 500
                        ? " The endpoint failed internally. Check its server logs."
                        : string.Empty;
            }
        }
    }
}

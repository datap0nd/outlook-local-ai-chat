using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookLocalAIChat;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.UI;

namespace GuardrailTests
{
    internal static class Program
    {
        private static int _passed;

        private static int Main()
        {
            try
            {
                Run("HTTPS endpoint is accepted", HttpsEndpointIsAccepted);
                Run("Loopback HTTP endpoint is accepted", LoopbackHttpIsAccepted);
                Run("Remote HTTP endpoint is rejected", RemoteHttpIsRejected);
                Run(
                    "Remote HTTP endpoint requires explicit opt in",
                    RemoteHttpIsAcceptedWithOptIn);
                Run(
                    "Models endpoint is normalized",
                    ModelsEndpointIsNormalized);
                Run(
                    "Recommended model balances capability and speed",
                    RecommendedModelIsStable);
                Run(
                    "Model selection never recommends Gauss variants",
                    GaussModelsAreNeverRecommended);
                Run(
                    "Compatible endpoint model discovery is verified",
                    ModelDiscoveryUsesCompatibleContract);
                Run(
                    "Compatible tool calls tolerate a missing call id",
                    ToolCallResponseIsNormalized);
                Run("Text boundary removes controls and truncates", TextIsBounded);
                Run(
                    "Endpoint diagnostics expose transport details",
                    EndpointDiagnosticsExposeTransportDetails);
                Run("Mailbox tools are read only", MailboxToolsAreReadOnly);
                Run("Request schema exposes bounded tools", RequestSchemaIsBounded);
                Run("Email is labeled as untrusted data", EmailIsUntrustedData);
                Run("Conversation history is bounded", HistoryIsBounded);
                Run("Draft service exposes no send capability", DraftHasNoSend);
                Run(
                    "Mailbox host exposes one guarded dispatcher",
                    MailboxHostHasGuardedDispatcher);
                Run(
                    "Office startup and task pane COM interfaces are dual",
                    OfficeStartupInterfacesAreDual);
                Run(
                    "Chat pane is a registered COM control",
                    ChatPaneIsComControl);
                Console.WriteLine("PASS: " + _passed + " guardrail tests");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("FAIL: " + exception.Message);
                return 1;
            }
        }

        private static void Run(string name, Action test)
        {
            test();
            _passed++;
            Console.WriteLine("PASS: " + name);
        }

        private static void HttpsEndpointIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "https://ai.example.test/v1",
                    out endpoint),
                "HTTPS should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/chat/completions",
                "The chat completions path was not normalized.");
        }

        private static void LoopbackHttpIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://127.0.0.1:1234/v1",
                    out endpoint),
                "Loopback HTTP should be accepted.");
        }

        private static void RemoteHttpIsRejected()
        {
            Uri endpoint;
            Assert(
                !AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1",
                    out endpoint),
                "Remote HTTP must be rejected.");
        }

        private static void RemoteHttpIsAcceptedWithOptIn()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1/chat/completions",
                    true,
                    out endpoint),
                "Explicitly allowed remote HTTP should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "http://ai.example.test/v1/chat/completions",
                "Remote HTTP endpoint was not normalized correctly.");

            var settings = new AppSettings
            {
                BaseUrl = "http://ai.example.test/v1",
                Model = "local-model",
                ApiKey = "test-key",
                AllowInsecureHttp = true
            };
            Assert(
                settings.IsConfigured,
                "Remote HTTP opt in was not honored by configuration.");
        }

        private static void ModelsEndpointIsNormalized()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetModelsUri(
                    "https://ai.example.test/v1/chat/completions",
                    false,
                    out endpoint),
                "Models endpoint should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/models",
                "The models path was not normalized.");
        }

        private static void RecommendedModelIsStable()
        {
            Assert(
                ModelSelectionPolicy.RecommendedModel ==
                "qwen3.5-35b-a3b",
                "The balanced default changed unexpectedly.");
            var chosen = ModelSelectionPolicy.ChooseRecommended(
                new[]
                {
                    "gpt-oss-120b",
                    "qwen3.5-35b-a3b",
                    "gpt-oss-20b"
                });
            Assert(
                chosen == "qwen3.5-35b-a3b",
                "The balanced model was not preferred: " + chosen);
        }

        private static void GaussModelsAreNeverRecommended()
        {
            var chosen = ModelSelectionPolicy.ChooseRecommended(
                new[]
                {
                    "gausso",
                    "gausso-flash",
                    "gauss-think",
                    "gpt-oss-20b"
                });
            Assert(
                chosen == "gpt-oss-20b",
                "A Gauss variant was recommended: " + chosen);
            Assert(
                ModelSelectionPolicy.ChooseRecommended(
                    new[] { "gausso", "gauss" }) ==
                string.Empty,
                "Gauss-only endpoints must not produce a recommendation.");
        }

        private static void ModelDiscoveryUsesCompatibleContract()
        {
            const string response =
                "{\"data\":[" +
                "{\"id\":\"qwen3.5-35b-a3b\"}," +
                "{\"id\":\"gpt-oss-20b\"}," +
                "{\"id\":\"text-embedding-model\"}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var models = client.GetModelsAsync(
                    settings,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine == "GET /v1/models HTTP/1.1",
                    "Unexpected model-list request: " +
                    server.RequestLine);
                Assert(
                    server.Authorization == "Bearer test-key",
                    "The API key was not sent as a Bearer token.");
                Assert(
                    models.SequenceEqual(
                        new[]
                        {
                            "gpt-oss-20b",
                            "qwen3.5-35b-a3b"
                        }),
                    "Unexpected discovered models: " +
                    string.Join(", ", models));
            }
        }

        private static void ToolCallResponseIsNormalized()
        {
            const string response =
                "{\"choices\":[{\"message\":{" +
                "\"role\":\"assistant\",\"content\":null," +
                "\"tool_calls\":[{\"function\":{" +
                "\"name\":\"search_mailbox\"," +
                "\"arguments\":\"{}\"}}]}}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var request = MakeRequest(
                    new List<ChatTurn>());
                request.model = settings.Model;
                var message = client.CompleteAsync(
                    settings,
                    request,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine ==
                    "POST /v1/chat/completions HTTP/1.1",
                    "Unexpected completion request: " +
                    server.RequestLine);
                Assert(
                    server.Body.Contains(
                        "\"model\":\"qwen3.5-35b-a3b\"") &&
                    server.Body.Contains("\"stream\":false") &&
                    server.Body.Contains("\"tools\""),
                    "The completion request contract is incomplete.");
                Assert(
                    message.tool_calls.Count == 1 &&
                    message.tool_calls[0].id == "call_1" &&
                    message.tool_calls[0].type == "function",
                    "The missing tool-call identity was not normalized.");
            }
        }

        private static AppSettings EndpointSettings(
            string baseUrl)
        {
            return new AppSettings
            {
                BaseUrl = baseUrl,
                Model =
                    ModelSelectionPolicy.RecommendedModel,
                ApiKey = "test-key"
            };
        }

        private static void TextIsBounded()
        {
            var result = TextBoundary.PlainText("a\u0000bcd", 3);
            Assert(result == "abc", "Unexpected bounded text: " + result);
        }

        private static void EndpointDiagnosticsExposeTransportDetails()
        {
            var exception = new AiEndpointException(
                "NETWORK_CONNECT_FAILURE",
                "The endpoint could not be reached.",
                transportDetails:
                    "SocketError ConnectionRefused NativeError 10061");
            var diagnostic = exception.ToDiagnosticText();
            Assert(
                diagnostic.Contains("[NETWORK_CONNECT_FAILURE]") &&
                diagnostic.Contains("Transport details:") &&
                diagnostic.Contains("ConnectionRefused") &&
                diagnostic.Contains("10061"),
                "Transport diagnostics are incomplete: " + diagnostic);
        }

        private static void MailboxToolsAreReadOnly()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var json = new JavaScriptSerializer().Serialize(request);
            Assert(
                json.Contains("\"tools\"") &&
                json.Contains("\"tool_choice\":\"auto\""),
                "Request does not expose bounded mailbox tools.");
            Assert(json.Contains("\"stream\":false"), "Streaming must be off.");

            var names = request.tools
                .Select(tool => tool.function.name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "read_messages",
                "read_thread",
                "search_mailbox"
            };
            Assert(
                names.SequenceEqual(expected),
                "Unexpected mailbox tools: " +
                string.Join(", ", names));
        }

        private static void RequestSchemaIsBounded()
        {
            var fields = typeof(ChatCompletionRequest)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(property => property.Name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "messages",
                "model",
                "stream",
                "tool_choice",
                "tools"
            };
            Assert(
                fields.SequenceEqual(expected),
                "Model request capabilities changed: " +
                string.Join(", ", fields));
        }

        private static void EmailIsUntrustedData()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var context =
                ((ChatCompletionInputMessage)request.messages[1])
                .content;
            Assert(
                context.Contains("<selected_email_reference") &&
                context.Contains("untrusted reference data") &&
                !context.Contains("Message body"),
                "Email boundary markers are missing.");
        }

        private static void HistoryIsBounded()
        {
            var history = Enumerable.Range(0, 30)
                .Select(index => new ChatTurn("user", "turn " + index))
                .ToList();
            var request = MakeRequest(history);
            var historyMessages = request.messages.Count - 3;
            Assert(
                historyMessages == TextBoundary.MaxConversationTurns,
                "Unexpected retained history count: " + historyMessages);
        }

        private static void DraftHasNoSend()
        {
            var methods = typeof(DraftService)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method => method.DeclaringType == typeof(DraftService))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Length == 2 &&
                methods.Contains("CreateReplyDraft") &&
                methods.Contains("CreateNewDraft"),
                "Draft service public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void MailboxHostHasGuardedDispatcher()
        {
            var methods = typeof(MailboxToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method =>
                    method.DeclaringType ==
                    typeof(MailboxToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Length == 1 &&
                methods[0] == "Execute",
                "Mailbox tool host public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void OfficeStartupInterfacesAreDual()
        {
            Assert(
                typeof(IDTExtensibility2).IsImport,
                "IDTExtensibility2 must be a COM-import interface.");
            Assert(
                typeof(IRibbonExtensibility).IsImport,
                "IRibbonExtensibility must be a COM-import interface.");
            Assert(
                typeof(ICustomTaskPaneConsumer).IsImport,
                "ICustomTaskPaneConsumer must be a COM-import interface.");

            AssertDual(typeof(IDTExtensibility2));
            AssertDual(typeof(IRibbonExtensibility));
            AssertDual(typeof(ICustomTaskPaneConsumer));

            var addIn = new AddIn();
            AssertComInterface(addIn, typeof(IDTExtensibility2));
            AssertComInterface(addIn, typeof(IRibbonExtensibility));
            AssertComInterface(
                addIn,
                typeof(ICustomTaskPaneConsumer));
        }

        private static void ChatPaneIsComControl()
        {
            var type = typeof(ChatPane);
            var visible = type
                .GetCustomAttributes(
                    typeof(ComVisibleAttribute),
                    false)
                .Cast<ComVisibleAttribute>()
                .Single();
            var progId = type
                .GetCustomAttributes(
                    typeof(ProgIdAttribute),
                    false)
                .Cast<ProgIdAttribute>()
                .Single();
            Assert(visible.Value, "ChatPane must be COM visible.");
            Assert(
                progId.Value ==
                "OutlookLocalAIChat.ChatPane",
                "Unexpected ChatPane ProgID.");
            Assert(
                type.GUID ==
                new Guid(
                    "14D24FA1-4342-442F-B68B-B68D7372794C"),
                "Unexpected ChatPane CLSID.");
        }

        private static void AssertDual(Type interfaceType)
        {
            var attribute = interfaceType
                .GetCustomAttributes(typeof(TypeLibTypeAttribute), false)
                .Cast<TypeLibTypeAttribute>()
                .Single();
            var expected =
                TypeLibTypeFlags.FDispatchable |
                TypeLibTypeFlags.FDual;
            Assert(
                (attribute.Value & expected) == expected,
                interfaceType.Name + " must be a dual dispatch interface.");
        }

        private static void AssertComInterface(object instance, Type interfaceType)
        {
            var pointer = Marshal.GetComInterfaceForObject(
                instance,
                interfaceType);
            try
            {
                Assert(
                    pointer != IntPtr.Zero,
                    interfaceType.Name + " was not exposed by the add-in.");
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    Marshal.Release(pointer);
                }
            }
        }

        private static ChatCompletionRequest MakeRequest(
            IReadOnlyList<ChatTurn> history)
        {
            return ChatRequestFactory.Create(
                "local-model",
                new MessageSnapshot(
                    "entry",
                    "store",
                    "Subject",
                    "Sender",
                    "Recipient",
                    DateTime.UtcNow,
                    "Message body"),
                history,
                "Help me reply.");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private sealed class FakeEndpoint : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly Task _requestTask;
            private readonly string _responseBody;

            public FakeEndpoint(string responseBody)
            {
                _responseBody = responseBody;
                _listener = new TcpListener(
                    IPAddress.Loopback,
                    0);
                _listener.Start();
                var port =
                    ((IPEndPoint)_listener.LocalEndpoint)
                    .Port;
                BaseUrl = "http://127.0.0.1:" +
                    port + "/v1";
                _requestTask = Task.Run(
                    (Action)HandleRequest);
            }

            public string BaseUrl { get; }

            public string RequestLine { get; private set; } =
                string.Empty;

            public string Authorization { get; private set; } =
                string.Empty;

            public string Body { get; private set; } =
                string.Empty;

            public void Wait()
            {
                if (!_requestTask.Wait(
                    TimeSpan.FromSeconds(10)))
                {
                    throw new InvalidOperationException(
                        "The fake endpoint did not receive a request.");
                }

                if (_requestTask.IsFaulted)
                {
                    throw _requestTask.Exception
                        .GetBaseException();
                }
            }

            public void Dispose()
            {
                _listener.Stop();
                try
                {
                    _requestTask.Wait(
                        TimeSpan.FromSeconds(1));
                }
                catch
                {
                }
            }

            private void HandleRequest()
            {
                using (var client = _listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(
                    stream,
                    Encoding.ASCII,
                    false,
                    4096,
                    true))
                {
                    RequestLine =
                        reader.ReadLine() ?? string.Empty;
                    var contentLength = 0;
                    while (true)
                    {
                        var line = reader.ReadLine();
                        if (string.IsNullOrEmpty(line))
                        {
                            break;
                        }

                        var separator = line.IndexOf(':');
                        if (separator <= 0)
                        {
                            continue;
                        }

                        var name = line
                            .Substring(0, separator)
                            .Trim();
                        var value = line
                            .Substring(separator + 1)
                            .Trim();
                        if (name.Equals(
                            "Authorization",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Authorization = value;
                        }
                        else if (name.Equals(
                            "Content-Length",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(
                                value,
                                out contentLength);
                        }
                    }

                    if (contentLength > 0)
                    {
                        var buffer =
                            new char[contentLength];
                        var offset = 0;
                        while (offset < contentLength)
                        {
                            var read = reader.Read(
                                buffer,
                                offset,
                                contentLength - offset);
                            if (read <= 0)
                            {
                                break;
                            }

                            offset += read;
                        }

                        Body = new string(
                            buffer,
                            0,
                            offset);
                    }

                    var responseBytes =
                        Encoding.UTF8.GetBytes(
                            _responseBody);
                    var headers = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\n" +
                        "Content-Type: application/json\r\n" +
                        "Content-Length: " +
                        responseBytes.Length +
                        "\r\nConnection: close\r\n\r\n");
                    stream.Write(
                        headers,
                        0,
                        headers.Length);
                    stream.Write(
                        responseBytes,
                        0,
                        responseBytes.Length);
                    stream.Flush();
                }
            }
        }
    }
}

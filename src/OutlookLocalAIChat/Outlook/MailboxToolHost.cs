using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class MailboxToolHost
    {
        private const int MaxDirectMessageBodyCharacters = 6000;
        private const int MaxThreadMessageBodyCharacters = 2000;
        private const int MaxThreadMessages = 12;

        private readonly MailboxContextService _mailbox;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly Dictionary<string, MessageSnapshot> _handles =
            new Dictionary<string, MessageSnapshot>(
                StringComparer.Ordinal);
        private int _nextHandle = 1;

        public MailboxToolHost(
            object outlookApplication,
            MessageSnapshot selectedMessage)
        {
            _mailbox = new MailboxContextService(outlookApplication);
            if (selectedMessage != null)
            {
                _handles["selected"] = selectedMessage;
            }
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id))
            {
                return Error(
                    call?.id,
                    "MAILBOX_TOOL_CALL_INVALID",
                    "The model returned an invalid tool call.");
            }

            var name = call.function.name ?? string.Empty;
            if (!MailboxToolCatalog.IsApproved(name))
            {
                return Error(
                    call.id,
                    "MAILBOX_TOOL_NOT_ALLOWED",
                    "The requested mailbox tool is not allowed.");
            }

            try
            {
                var arguments = ParseArguments(
                    call.function.arguments);
                switch (name)
                {
                    case MailboxToolCatalog.SearchMailbox:
                        return Search(call.id, arguments);
                    case MailboxToolCatalog.ReadMessages:
                        return ReadMessages(call.id, arguments);
                    case MailboxToolCatalog.ReadThread:
                        return ReadThread(call.id, arguments);
                    default:
                        return Error(
                            call.id,
                            "MAILBOX_TOOL_NOT_ALLOWED",
                            "The requested mailbox tool is not allowed.");
                }
            }
            catch (Exception exception)
            {
                Log.Error("MailboxTool." + name, exception);
                return Error(
                    call.id,
                    "MAILBOX_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "MAILBOX_TOOL_FAILED"));
            }
        }

        internal MessageSnapshot ResolveHandle(string handle)
        {
            var boundedHandle = TextBoundary.SingleLine(
                handle,
                64);
            MessageSnapshot message;
            return _handles.TryGetValue(boundedHandle, out message)
                ? message
                : null;
        }

        private MailboxToolResult Search(
            string callId,
            IDictionary<string, object> arguments)
        {
            var query = GetString(arguments, "query", string.Empty);
            var folder = GetString(arguments, "folder", "all")
                .ToLowerInvariant();
            if (folder != "all" &&
                folder != "inbox" &&
                folder != "sent")
            {
                folder = "all";
            }

            var daysBack = GetInteger(
                arguments,
                "days_back",
                365,
                1,
                3650);
            var maxResults = GetInteger(
                arguments,
                "max_results",
                10,
                1,
                20);
            var hits = _mailbox.Search(
                query,
                folder,
                daysBack,
                maxResults);

            var results = new List<object>();
            foreach (var hit in hits)
            {
                var handle = Register(hit.Message);
                results.Add(new Dictionary<string, object>
                {
                    { "handle", handle },
                    { "folder", hit.FolderName },
                    { "subject", hit.Message.Subject },
                    { "from", hit.Message.Sender },
                    { "to", hit.Message.Recipients },
                    {
                        "received",
                        hit.Message.ReceivedAt?.ToString("O") ??
                        "unknown"
                    },
                    { "snippet", hit.Snippet }
                });
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "query", query },
                    { "folder", folder },
                    { "result_count", results.Count },
                    { "results", results }
                },
                "Mailbox search loaded " +
                results.Count.ToString(CultureInfo.InvariantCulture) +
                " result summaries.");
        }

        private MailboxToolResult ReadMessages(
            string callId,
            IDictionary<string, object> arguments)
        {
            var handles = GetStringList(arguments, "handles")
                .Distinct(StringComparer.Ordinal)
                .Take(4)
                .ToArray();
            if (handles.Length == 0)
            {
                return Error(
                    callId,
                    "MAILBOX_HANDLES_REQUIRED",
                    "At least one message handle is required.");
            }

            var messages = new List<object>();
            foreach (var handle in handles)
            {
                MessageSnapshot message;
                if (!_handles.TryGetValue(handle, out message))
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "MAILBOX_HANDLE_UNKNOWN" }
                    });
                    continue;
                }

                messages.Add(
                    SerializeMessage(
                        handle,
                        message,
                        MaxDirectMessageBodyCharacters));
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "messages", messages }
                },
                "Loaded " +
                messages.Count.ToString(CultureInfo.InvariantCulture) +
                " message bodies.");
        }

        private MailboxToolResult ReadThread(
            string callId,
            IDictionary<string, object> arguments)
        {
            var handle = GetString(arguments, "handle", string.Empty);
            MessageSnapshot source;
            if (!_handles.TryGetValue(handle, out source))
            {
                return Error(
                    callId,
                    "MAILBOX_HANDLE_UNKNOWN",
                    "The message handle is unknown or expired.");
            }

            var conversation = _mailbox.ReadConversation(
                source,
                MaxThreadMessages);
            var messages = new List<object>();
            foreach (var message in conversation)
            {
                var messageHandle = Register(message);
                messages.Add(
                    SerializeMessage(
                        messageHandle,
                        message,
                        MaxThreadMessageBodyCharacters));
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "source_handle", handle },
                    { "message_count", messages.Count },
                    { "messages", messages }
                },
                "Loaded " +
                messages.Count.ToString(CultureInfo.InvariantCulture) +
                " messages from the conversation.");
        }

        private object SerializeMessage(
            string handle,
            MessageSnapshot message,
            int maximumBodyCharacters)
        {
            return new Dictionary<string, object>
            {
                { "handle", handle },
                { "subject", message.Subject },
                { "from", message.Sender },
                { "to", message.Recipients },
                {
                    "received",
                    message.ReceivedAt?.ToString("O") ??
                    "unknown"
                },
                {
                    "body",
                    TextBoundary.PlainText(
                        message.Body,
                        maximumBodyCharacters)
                }
            };
        }

        private string Register(MessageSnapshot message)
        {
            foreach (var pair in _handles)
            {
                if (pair.Value.EntryId.Equals(
                        message.EntryId,
                        StringComparison.Ordinal) &&
                    pair.Value.StoreId.Equals(
                        message.StoreId,
                        StringComparison.Ordinal))
                {
                    return pair.Key;
                }
            }

            var handle = "m" +
                _nextHandle.ToString(CultureInfo.InvariantCulture);
            _nextHandle++;
            _handles[handle] = message;
            return handle;
        }

        private MailboxToolResult Success(
            string callId,
            object payload,
            string status)
        {
            var json = _serializer.Serialize(payload);
            if (json.Length > TextBoundary.MaxToolResultCharacters)
            {
                return Error(
                    callId,
                    "MAILBOX_TOOL_RESULT_TOO_LARGE",
                    "The bounded mailbox result was still too large to return safely.");
            }

            return new MailboxToolResult(
                callId,
                json,
                status);
        }

        private MailboxToolResult Error(
            string callId,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        }
                    }),
                "[" + code + "] " + message);
        }

        private IDictionary<string, object> ParseArguments(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            try
            {
                var parsed = _serializer.DeserializeObject(json) as
                    IDictionary<string, object>;
                if (parsed == null)
                {
                    throw new InvalidOperationException(
                        "Tool arguments must be a JSON object.");
                }

                return parsed;
            }
            catch (Exception exception)
            {
                throw new AiEndpointException(
                    "TOOL_ARGUMENTS_INVALID_JSON",
                    "The model returned invalid JSON tool arguments.",
                    exception,
                    responseSnippet: json);
            }
        }

        private static string GetString(
            IDictionary<string, object> arguments,
            string key,
            string fallback)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.PlainText(
                    Convert.ToString(value),
                    1000)
                : fallback;
        }

        private static int GetInteger(
            IDictionary<string, object> arguments,
            string key,
            int fallback,
            int minimum,
            int maximum)
        {
            object value;
            int parsed;
            if (!arguments.TryGetValue(key, out value) ||
                !int.TryParse(
                    Convert.ToString(value),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return fallback;
            }

            return Math.Max(minimum, Math.Min(maximum, parsed));
        }

        private static IEnumerable<string> GetStringList(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            if (!arguments.TryGetValue(key, out value))
            {
                return new string[0];
            }

            var array = value as object[];
            if (array != null)
            {
                return array
                    .Select(item =>
                        TextBoundary.PlainText(
                            Convert.ToString(item),
                            100))
                    .Where(item => item.Length > 0);
            }

            var list = value as ArrayList;
            if (list != null)
            {
                return list
                    .Cast<object>()
                    .Select(item =>
                        TextBoundary.PlainText(
                            Convert.ToString(item),
                            100))
                    .Where(item => item.Length > 0);
            }

            return new[]
            {
                TextBoundary.PlainText(
                    Convert.ToString(value),
                    100)
            }.Where(item => item.Length > 0);
        }
    }
}

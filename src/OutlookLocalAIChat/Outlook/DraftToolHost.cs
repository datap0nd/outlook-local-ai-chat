using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class DraftToolHost
    {
        private readonly object _outlookApplication;
        private readonly MessageSnapshot _selectedMessage;
        private readonly OneShotDraftAuthorization _authorization;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public DraftToolHost(
            object outlookApplication,
            MessageSnapshot selectedMessage,
            OneShotDraftAuthorization authorization)
        {
            _outlookApplication = outlookApplication;
            _selectedMessage = selectedMessage;
            _authorization = authorization ??
                throw new ArgumentNullException(nameof(authorization));
        }

        public MailboxToolResult Execute(
            ChatToolCall call,
            bool isOnlyToolCall)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !DraftToolCatalog.IsCreateDraft(
                    call.function.name))
            {
                return Error(
                    call?.id,
                    "DRAFT_TOOL_CALL_INVALID",
                    "The model returned an invalid draft tool call.");
            }

            if (!isOnlyToolCall)
            {
                return Error(
                    call.id,
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
                    "create_draft must be the only tool call in its response.");
            }

            IDictionary<string, object> arguments;
            try
            {
                arguments = ParseArguments(
                    call.function.arguments);
            }
            catch (Exception exception)
            {
                return Error(
                    call.id,
                    "DRAFT_ARGUMENTS_INVALID_JSON",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_ARGUMENTS_INVALID_JSON"));
            }

            var kind = GetString(arguments, "kind")
                .ToLowerInvariant();
            var body = TextBoundary.PlainText(
                GetString(arguments, "body"),
                TextBoundary.MaxAssistantCharacters);
            if (kind != "new" && kind != "reply")
            {
                return Error(
                    call.id,
                    "DRAFT_KIND_INVALID",
                    "Draft kind must be new or reply.");
            }

            if (body.Length == 0)
            {
                return Error(
                    call.id,
                    "DRAFT_BODY_REQUIRED",
                    "A non-empty plain-text draft body is required.");
            }

            if (kind == "reply" &&
                (_selectedMessage == null ||
                 !_selectedMessage.CanReply))
            {
                return Error(
                    call.id,
                    "DRAFT_REPLY_SOURCE_REQUIRED",
                    "Select a reply-capable email before authorizing a reply draft.");
            }

            if (!_authorization.TryConsume())
            {
                return Error(
                    call.id,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "No unused one-shot draft permission is available for this request.");
            }

            try
            {
                var drafts = new DraftService(
                    _outlookApplication);
                if (kind == "reply")
                {
                    drafts.CreateReplyDraft(
                        _selectedMessage,
                        body);
                }
                else
                {
                    drafts.CreateNewDraft(
                        body,
                        GetString(arguments, "subject"),
                        GetString(arguments, "to"),
                        GetString(arguments, "cc"));
                }

                _authorization.MarkCreated();
                return Success(
                    call.id,
                    kind);
            }
            catch (Exception exception)
            {
                Log.Error("DraftTool.CreateDraft", exception);
                return Error(
                    call.id,
                    "DRAFT_CREATION_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_CREATION_FAILED"));
            }
        }

        private IDictionary<string, object> ParseArguments(
            string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            var parsed = _serializer.DeserializeObject(json) as
                IDictionary<string, object>;
            if (parsed == null)
            {
                throw new InvalidOperationException(
                    "Draft tool arguments must be a JSON object.");
            }

            return parsed;
        }

        private static string GetString(
            IDictionary<string, object> arguments,
            string name)
        {
            object value;
            return arguments != null &&
                   arguments.TryGetValue(name, out value)
                ? Convert.ToString(value) ?? string.Empty
                : string.Empty;
        }

        private MailboxToolResult Success(
            string callId,
            string kind)
        {
            return new MailboxToolResult(
                callId,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", true },
                        { "draft_kind", kind },
                        { "sent", false },
                        { "permission_consumed", true }
                    }),
                kind == "reply"
                    ? "One unsent reply draft opened in Outlook."
                    : "One unsent new-message draft opened in Outlook.");
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
                        },
                        {
                            "permission_consumed",
                            _authorization.IsConsumed
                        }
                    }),
                "[" + code + "] " +
                TextBoundary.PlainText(message, 600));
        }
    }
}

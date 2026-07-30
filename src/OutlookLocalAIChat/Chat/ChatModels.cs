using System.Collections.Generic;

namespace OutlookLocalAIChat.Chat
{
    public sealed class ChatTurn
    {
        public ChatTurn(string role, string content)
        {
            Role = role;
            Content = content;
        }

        public string Role { get; }

        public string Content { get; }
    }

    public sealed class ChatCompletionRequest
    {
        public string model { get; set; }

        public List<object> messages { get; set; }

        public bool stream { get; set; }

        public List<ChatToolDefinition> tools { get; set; }

        public string tool_choice { get; set; }
    }

    public sealed class ChatCompletionInputMessage
    {
        public string role { get; set; }

        public string content { get; set; }
    }

    public sealed class ChatCompletionAssistantToolMessage
    {
        public string role { get; set; }

        public string content { get; set; }

        public List<ChatToolCall> tool_calls { get; set; }
    }

    public sealed class ChatCompletionToolResultMessage
    {
        public string role { get; set; }

        public string tool_call_id { get; set; }

        public string content { get; set; }
    }

    public sealed class ChatCompletionResponse
    {
        public List<ChatCompletionChoice> choices { get; set; }

        public ChatCompletionError error { get; set; }
    }

    public sealed class ModelListResponse
    {
        public List<ModelListItem> data { get; set; }
    }

    public sealed class ModelListItem
    {
        public string id { get; set; }
    }

    public sealed class ChatCompletionChoice
    {
        public ChatCompletionResponseMessage message { get; set; }
    }

    public sealed class ChatCompletionResponseMessage
    {
        public string role { get; set; }

        public string content { get; set; }

        public List<ChatToolCall> tool_calls { get; set; }
    }

    public sealed class ChatCompletionError
    {
        public string message { get; set; }

        public string type { get; set; }

        public string code { get; set; }
    }

    public sealed class ChatToolDefinition
    {
        public string type { get; set; }

        public ChatToolFunctionDefinition function { get; set; }
    }

    public sealed class ChatToolFunctionDefinition
    {
        public string name { get; set; }

        public string description { get; set; }

        public object parameters { get; set; }
    }

    public sealed class ChatToolCall
    {
        public string id { get; set; }

        public string type { get; set; }

        public ChatToolCallFunction function { get; set; }
    }

    public sealed class ChatToolCallFunction
    {
        public string name { get; set; }

        public string arguments { get; set; }
    }

    public sealed class MailboxToolResult
    {
        public MailboxToolResult(
            string toolCallId,
            string content,
            string statusText)
        {
            ToolCallId = toolCallId ?? string.Empty;
            Content = content ?? string.Empty;
            StatusText = statusText ?? string.Empty;
        }

        public string ToolCallId { get; }

        public string Content { get; }

        public string StatusText { get; }
    }
}

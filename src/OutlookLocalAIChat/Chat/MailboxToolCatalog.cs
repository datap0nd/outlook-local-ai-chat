using System;
using System.Collections.Generic;

namespace OutlookLocalAIChat.Chat
{
    public static class MailboxToolCatalog
    {
        public const string SearchMailbox = "search_mailbox";
        public const string ReadMessages = "read_messages";
        public const string ReadThread = "read_thread";

        public static readonly IReadOnlyList<string> ApprovedNames =
            new[]
            {
                SearchMailbox,
                ReadMessages,
                ReadThread
            };

        public static List<ChatToolDefinition> CreateDefinitions(
            bool workingSetOnly = false)
        {
            var definitions = new List<ChatToolDefinition>
            {
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SearchMailbox,
                        description =
                            "Search the user's local primary Outlook Inbox and Sent Items. " +
                            "Returns bounded metadata, snippets, and temporary message handles.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "query",
                                    StringSchema(
                                        "Words or a phrase to find in subject, body, sender, or recipients. " +
                                        "Use an empty string to list recent mail.")
                                },
                                {
                                    "folder",
                                    EnumSchema(
                                        "Mailbox folders to search.",
                                        "all",
                                        "inbox",
                                        "sent")
                                },
                                {
                                    "days_back",
                                    IntegerSchema(
                                        "Only include messages this many days old or newer.",
                                        1,
                                        3650)
                                },
                                {
                                    "max_results",
                                    IntegerSchema(
                                        "Maximum number of result summaries to return.",
                                        1,
                                        5)
                                }
                            },
                            "query")
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ReadMessages,
                        description =
                            "Load the bounded plain-text bodies of messages returned by " +
                            "search_mailbox, the selected email, or the user-approved " +
                            "five-message working set. At most five unique message bodies " +
                            "can be loaded in one request.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "handles",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "array" },
                                        {
                                            "items",
                                            StringSchema(
                                                "A temporary handle returned by search_mailbox, " +
                                                "selected, or context1 through context5 from the " +
                                                "user-approved working set.")
                                        },
                                        { "minItems", 1 },
                                        { "maxItems", 5 }
                                    }
                                }
                            },
                            "handles")
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ReadThread,
                        description =
                            "Load bounded messages in the Outlook conversation containing " +
                            "a searched or selected message, subject to the five-message " +
                            "request-wide context cap.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "handle",
                                    StringSchema(
                                        "A temporary handle returned by search_mailbox, " +
                                        "or selected for the currently selected email.")
                                }
                            },
                            "handle")
                    }
                }
            };

            return workingSetOnly
                ? new List<ChatToolDefinition>
                {
                    definitions[1]
                }
                : definitions;
        }

        public static bool IsApproved(string name)
        {
            foreach (var approved in ApprovedNames)
            {
                if (string.Equals(
                    approved,
                    name,
                    StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, object> ObjectSchema(
            Dictionary<string, object> properties,
            params string[] required)
        {
            return new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", properties },
                { "required", required },
                { "additionalProperties", false }
            };
        }

        private static Dictionary<string, object> StringSchema(
            string description)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description }
            };
        }

        private static Dictionary<string, object> EnumSchema(
            string description,
            params string[] values)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description },
                { "enum", values }
            };
        }

        private static Dictionary<string, object> IntegerSchema(
            string description,
            int minimum,
            int maximum)
        {
            return new Dictionary<string, object>
            {
                { "type", "integer" },
                { "description", description },
                { "minimum", minimum },
                { "maximum", maximum }
            };
        }
    }
}

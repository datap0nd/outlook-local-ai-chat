using System.Collections.Generic;

namespace OutlookLocalAIChat.Chat
{
    public static class DraftToolCatalog
    {
        public const string CreateDraft = "create_draft";

        public static ChatToolDefinition CreateDefinition()
        {
            return new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = CreateDraft,
                    description =
                        "Create, save, and display exactly one unsent Outlook draft. " +
                        "This tool is available only after explicit one-request user " +
                        "authorization. It must be the only tool call in its response " +
                        "and can never send email.",
                    parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        {
                            "properties",
                            new Dictionary<string, object>
                            {
                                {
                                    "kind",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "string" },
                                        {
                                            "description",
                                            "Use reply for the selected email or new for a new message."
                                        },
                                        {
                                            "enum",
                                            new[] { "new", "reply" }
                                        }
                                    }
                                },
                                {
                                    "body",
                                    BoundedString(
                                        "Plain-text draft body. Do not include commentary outside the draft.",
                                        12000)
                                },
                                {
                                    "subject",
                                    BoundedString(
                                        "Optional subject for a new draft. Ignored for replies.",
                                        255)
                                },
                                {
                                    "to",
                                    BoundedString(
                                        "Optional semicolon-separated recipients for a new draft. Ignored for replies.",
                                        2000)
                                },
                                {
                                    "cc",
                                    BoundedString(
                                        "Optional semicolon-separated CC recipients for a new draft. Ignored for replies.",
                                        2000)
                                }
                            }
                        },
                        {
                            "required",
                            new[] { "kind", "body" }
                        },
                        { "additionalProperties", false }
                    }
                }
            };
        }

        public static bool IsCreateDraft(string name)
        {
            return string.Equals(
                CreateDraft,
                name,
                System.StringComparison.Ordinal);
        }

        private static Dictionary<string, object> BoundedString(
            string description,
            int maximumLength)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description },
                { "maxLength", maximumLength }
            };
        }
    }
}

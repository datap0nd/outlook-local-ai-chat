using System;
using System.Text;

namespace OutlookLocalAIChat.Security
{
    public static class TextBoundary
    {
        public const int MaxMessageBodyCharacters = 24000;
        public const int MaxUserPromptCharacters = 4000;
        public const int MaxAssistantCharacters = 12000;
        public const int MaxToneProfileCharacters = 5000;
        public const int MaxConversationTurns = 12;
        public const int MaxToolResultCharacters = 48000;
        public const int MaxToolRounds = 4;
        public const int MaxToolCallsPerRound = 4;
        public const int MaxHttpResponseCharacters = 1048576;

        public static string PlainText(string value, int maximumLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var limit = Math.Max(0, maximumLength);
            var builder = new StringBuilder(Math.Min(value.Length, limit));

            foreach (var character in value)
            {
                if (builder.Length >= limit)
                {
                    break;
                }

                if (character == '\r' || character == '\n' || character == '\t')
                {
                    builder.Append(character);
                    continue;
                }

                if (!char.IsControl(character))
                {
                    builder.Append(character);
                }
            }

            return builder.ToString().Trim();
        }

        public static string SingleLine(
            string value,
            int maximumLength)
        {
            return PlainText(value, maximumLength)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ')
                .Trim();
        }
    }
}

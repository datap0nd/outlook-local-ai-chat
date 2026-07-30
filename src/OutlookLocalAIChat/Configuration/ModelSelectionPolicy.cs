using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookLocalAIChat.Configuration
{
    public static class ModelSelectionPolicy
    {
        public const string RecommendedModel = "qwen3.5-35b-a3b";

        private static readonly string[] PreferredModels =
        {
            RecommendedModel,
            "gpt-oss-120b",
            "gpt-oss-20b",
            "gemma-4-31b-it",
            "gemma-4-26b-a4b-it",
            "qwen3-vl-30b"
        };

        public static IReadOnlyList<string> Presets
        {
            get { return PreferredModels; }
        }

        public static bool IsGenerativeModel(string model)
        {
            var value = (model ?? string.Empty).Trim();
            return value.Length > 0 &&
                   value.IndexOf(
                       "embedding",
                       StringComparison.OrdinalIgnoreCase) < 0;
        }

        public static bool IsDisallowedRecommendation(string model)
        {
            return (model ?? string.Empty).IndexOf(
                "gauss",
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string ChooseRecommended(
            IEnumerable<string> availableModels)
        {
            var available = (availableModels ??
                Enumerable.Empty<string>())
                .Where(IsGenerativeModel)
                .Where(model => !IsDisallowedRecommendation(model))
                .Select(model => model.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var preferred in PreferredModels)
            {
                var exact = available.FirstOrDefault(
                    model => model.Equals(
                        preferred,
                        StringComparison.OrdinalIgnoreCase));
                if (exact != null)
                {
                    return exact;
                }
            }

            return available.FirstOrDefault() ?? string.Empty;
        }

        public static string DescriptionFor(string model)
        {
            var value = (model ?? string.Empty).Trim();
            if (value.Equals(
                RecommendedModel,
                StringComparison.OrdinalIgnoreCase))
            {
                return "Recommended balance of mailbox tool use, speed, and drafting quality.";
            }

            if (value.Equals(
                "gpt-oss-120b",
                StringComparison.OrdinalIgnoreCase))
            {
                return "Quality-first option. Expect higher latency and heavier server load.";
            }

            if (value.Equals(
                "gpt-oss-20b",
                StringComparison.OrdinalIgnoreCase))
            {
                return "Speed-first option. Useful when latency matters more than depth.";
            }

            if (IsDisallowedRecommendation(value))
            {
                return "Custom model. This add-in does not recommend Gauss or Gausso models.";
            }

            return "Custom model. It must support OpenAI-compatible chat tool calls.";
        }
    }
}

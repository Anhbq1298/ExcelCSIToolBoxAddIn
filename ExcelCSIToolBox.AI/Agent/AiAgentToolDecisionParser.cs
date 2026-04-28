using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelCSIToolBox.AI.Agent
{
    /// <summary>
    /// Parses the raw text returned by the tool-routing LLM into an AiAgentToolDecision.
    /// The LLM is asked to return JSON only. This parser is tolerant of extra surrounding text.
    /// </summary>
    public static class AiAgentToolDecisionParser
    {
        /// <summary>
        /// Attempt to parse the LLM response into a tool decision.
        /// Returns a "no tool call" decision if parsing fails.
        /// </summary>
        public static AiAgentToolDecision Parse(string llmResponse)
        {
            if (string.IsNullOrWhiteSpace(llmResponse))
            {
                return NoToolDecision("LLM returned an empty response.");
            }

            try
            {
                // Extract the first JSON object from the response (LLM sometimes adds extra text).
                string json = ExtractFirstJsonObject(llmResponse);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return NoToolDecision("Could not locate JSON in LLM response: " + llmResponse);
                }

                JObject obj = JObject.Parse(json);

                bool   shouldCallTool = obj.Value<bool>("shouldCallTool");
                string toolName       = obj.Value<string>("toolName") ?? string.Empty;
                string argumentsJson  = obj.Value<string>("argumentsJson") ?? "{}";
                string reason         = obj.Value<string>("reason") ?? string.Empty;

                return new AiAgentToolDecision
                {
                    ShouldCallTool = shouldCallTool,
                    ToolName       = toolName.Trim(),
                    ArgumentsJson  = string.IsNullOrWhiteSpace(argumentsJson) ? "{}" : argumentsJson,
                    Reason         = reason
                };
            }
            catch (Exception ex)
            {
                return NoToolDecision("Failed to parse LLM tool decision: " + ex.Message);
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────────

        private static AiAgentToolDecision NoToolDecision(string reason)
        {
            return new AiAgentToolDecision
            {
                ShouldCallTool = false,
                ToolName       = string.Empty,
                ArgumentsJson  = "{}",
                Reason         = reason
            };
        }

        /// <summary>Extract the first { ... } JSON block from the given text.</summary>
        private static string ExtractFirstJsonObject(string text)
        {
            int start = text.IndexOf('{');
            if (start < 0)
            {
                return null;
            }

            int depth = 0;
            for (int i = start; i < text.Length; i++)
            {
                if (text[i] == '{')  { depth++; }
                if (text[i] == '}')  { depth--; }

                if (depth == 0)
                {
                    return text.Substring(start, i - start + 1);
                }
            }

            return null;
        }
    }
}

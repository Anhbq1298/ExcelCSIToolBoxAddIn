using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class AgentTaskPlannerService
    {
        public IReadOnlyList<AgentTaskItem> CreateTasks(string userMessage)
        {
            var tasks = new List<AgentTaskItem>();
            foreach (string part in SplitIntoTaskTexts(userMessage))
            {
                var task = CreateTask(tasks.Count + 1, part);
                if (!string.IsNullOrWhiteSpace(task.OriginalText))
                {
                    tasks.Add(task);
                }
            }

            if (tasks.Count == 0 && !string.IsNullOrWhiteSpace(userMessage))
            {
                tasks.Add(CreateTask(1, userMessage));
            }

            return tasks;
        }

        private static IEnumerable<string> SplitIntoTaskTexts(string userMessage)
        {
            string text = (userMessage ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                yield break;
            }

            if (LooksLikeIntegratedTrussWorkflow(text))
            {
                yield return text;
                yield break;
            }

            string marked = Regex.Replace(
                text,
                @"\b(?:then|also|after\s+that|next|finally|rồi|roi|sau\s+đó|sau\s+do|thêm\s+nữa|them\s+nua|đồng\s+thời|dong\s+thoi)\b",
                "|",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            marked = Regex.Replace(
                marked,
                @"\s*(?:;|\r?\n)\s*",
                "|",
                RegexOptions.CultureInvariant);

            marked = Regex.Replace(
                marked,
                @"\s*,\s*(?=(?:add|create|draw|assign|apply|set|select|extract|get|return|summarize|check|read|update|fix|make|implement|generate|list|count|thêm|them|tạo|tao|gán|gan|chọn|chon|kiểm|kiem|đọc|doc|sửa|sua)\b)",
                "|",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            marked = Regex.Replace(
                marked,
                @"\s+\b(?:and|và|va)\b\s+(?=(?:add|create|draw|assign|apply|set|select|extract|get|return|summarize|check|read|update|fix|make|implement|generate|list|count|thêm|them|tạo|tao|gán|gan|chọn|chon|kiểm|kiem|đọc|doc|sửa|sua)\b)",
                "|",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            string[] pieces = marked.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string piece in pieces)
            {
                string cleaned = CleanupTaskText(piece);
                if (!string.IsNullOrWhiteSpace(cleaned))
                {
                    yield return cleaned;
                }
            }
        }

        private static string CleanupTaskText(string text)
        {
            string cleaned = (text ?? string.Empty).Trim();
            cleaned = Regex.Replace(cleaned, @"^(?:and|và|va|then|also|rồi|roi)\s+", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return cleaned.Trim(' ', '.', ',', ';');
        }

        private static AgentTaskItem CreateTask(int index, string originalText)
        {
            string normalized = NormalizeIntent(originalText);
            return new AgentTaskItem
            {
                Id = "task-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                OriginalText = originalText == null ? string.Empty : originalText.Trim(),
                NormalizedIntent = normalized,
                TargetObjectType = DetectTargetObjectType(normalized),
                ActionType = DetectActionType(normalized),
                Parameters = new Dictionary<string, string>(),
                Status = "Pending",
                ResultMessage = string.Empty
            };
        }

        private static string NormalizeIntent(string text)
        {
            return Regex.Replace((text ?? string.Empty).Trim().ToLowerInvariant(), @"\s+", " ");
        }

        private static string DetectTargetObjectType(string normalized)
        {
            if (ContainsAny(normalized, "point", "joint")) return "Point";
            if (ContainsAny(normalized, "frame", "beam", "member", "column", "brace")) return "Frame";
            if (ContainsAny(normalized, "shell", "area", "slab", "wall")) return "Shell";
            if (ContainsAny(normalized, "truss", "howe", "pratt", "warren", "mono-slope", "monoslope")) return "Truss";
            if (ContainsAny(normalized, "load pattern", "load combination", "load", "udl")) return "Load";
            if (ContainsAny(normalized, "section", "property", "prop")) return "Section";
            if (ContainsAny(normalized, "model", "unit", "units")) return "Model";
            return "General";
        }

        private static string DetectActionType(string normalized)
        {
            if (ContainsAny(normalized, "add", "create", "draw", "generate", "thêm", "them", "tạo", "tao")) return "Create";
            if (ContainsAny(normalized, "assign", "apply", "set", "gán", "gan")) return "Assign";
            if (ContainsAny(normalized, "select", "chọn", "chon")) return "Select";
            if (ContainsAny(normalized, "extract", "get", "list", "count", "check", "read", "return", "summarize")) return "Query";
            if (ContainsAny(normalized, "update", "fix", "implement", "make", "sửa", "sua")) return "CodeChange";
            return "Unknown";
        }

        private static bool ContainsAny(string text, params string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (text.IndexOf(values[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool LooksLikeIntegratedTrussWorkflow(string text)
        {
            string normalized = NormalizeIntent(text);
            bool hasTruss = ContainsAny(normalized, "truss", "howe", "pratt", "warren", "mono-slope", "monoslope");
            if (!hasTruss)
            {
                return false;
            }

            return ContainsAny(normalized, "udl", "distributed load", "top chord", "bottom chord", "span", "bays", "rising from");
        }
    }
}

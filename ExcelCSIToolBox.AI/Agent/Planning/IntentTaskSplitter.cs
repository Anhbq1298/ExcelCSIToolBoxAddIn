using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelCSIToolBox.AI.Agent
{
    public sealed class IntentTaskSplitter
    {
        public IEnumerable<string> Split(string userMessage)
        {
            string text = (userMessage ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                yield break;
            }

            string marked = Regex.Replace(
                text,
                @"\b(?:then|also|after\s+that|next|finally)\b",
                "|",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            marked = Regex.Replace(
                marked,
                @"\s+\b(?:and)\b\s+(?=(?:add|create|draw|assign|apply|set|select|delete|remove|update|modify|run|execute)\b)",
                "|",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            string[] pieces = marked.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < pieces.Length; i++)
            {
                string piece = pieces[i].Trim(' ', '.', ',', ';');
                if (!string.IsNullOrWhiteSpace(piece))
                {
                    yield return piece;
                }
            }
        }
    }
}

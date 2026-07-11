using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelCSIToolBox.Core.Models.CSI
{
    public sealed class CsiObjectIdentity
    {
        public string ObjectType { get; set; }

        public string UniqueName { get; set; }

        public string Label { get; set; }

        public string Story { get; set; }

        public IReadOnlyCollection<string> MatchKeys { get; set; }

        public bool Matches(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || MatchKeys == null)
            {
                return false;
            }

            string candidate = value.Trim();
            foreach (string key in MatchKeys)
            {
                if (string.Equals(key, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static CsiObjectIdentity Create(
            string objectType,
            string uniqueName,
            string label,
            string story,
            IEnumerable<string> additionalMatchKeys = null)
        {
            var keys = new List<string>();
            AddKey(keys, uniqueName);
            AddKey(keys, label);
            AddKey(keys, story);

            if (additionalMatchKeys != null)
            {
                foreach (string key in additionalMatchKeys)
                {
                    AddKey(keys, key);
                }
            }

            return new CsiObjectIdentity
            {
                ObjectType = CsiObjectTypes.Normalize(objectType),
                UniqueName = Clean(uniqueName),
                Label = Clean(label),
                Story = Clean(story),
                MatchKeys = new ReadOnlyCollection<string>(keys)
            };
        }

        private static void AddKey(ICollection<string> keys, string value)
        {
            string clean = Clean(value);
            if (string.IsNullOrWhiteSpace(clean))
            {
                return;
            }

            if (keys.Any(existing => string.Equals(existing, clean, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            keys.Add(clean);
        }

        private static string Clean(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }
}

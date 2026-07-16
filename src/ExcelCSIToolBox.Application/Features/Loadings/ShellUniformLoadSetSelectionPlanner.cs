using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.Features.Loadings
{
    public sealed class ShellUniformLoadSetSelectionPlan
    {
        public IReadOnlyList<string> RequestedLoadSetNames { get; set; } = new List<string>();

        public IReadOnlyList<string> MatchedLoadSetNames { get; set; } = new List<string>();

        public IReadOnlyList<string> UnknownLoadSetNames { get; set; } = new List<string>();

        public string SelectedStoryName { get; set; }

        public IReadOnlyList<string> SelectedStoryNames { get; set; } = new List<string>();

        public IReadOnlyList<string> AreaObjectNames { get; set; } = new List<string>();

        public IReadOnlyList<string> UnresolvedAreaReferences { get; set; } = new List<string>();

        public int MatchingAssignmentCount { get; set; }

        public int DuplicateShellCount { get; set; }

        public int UnresolvedAreaCount { get; set; }
    }

    public static class ShellUniformLoadSetSelectionPlanner
    {
        public static IReadOnlyList<string> NormalizeLoadSetNames(IEnumerable<string> loadSetNames)
        {
            List<string> names = CreateCanonicalNames(loadSetNames)
                .Values
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return names;
        }

        public static ShellUniformLoadSetSelectionPlan CreatePlan(
            IEnumerable<string> requestedLoadSetNames,
            IEnumerable<string> existingLoadSetNames,
            IEnumerable<ShellUniformLoadSetAreaAssignmentDto> assignments,
            Func<ShellUniformLoadSetAreaAssignmentDto, string> resolveAreaObjectName)
        {
            return CreatePlan(
                requestedLoadSetNames,
                existingLoadSetNames,
                assignments,
                resolveAreaObjectName,
                (IEnumerable<string>)null);
        }

        public static ShellUniformLoadSetSelectionPlan CreatePlan(
            IEnumerable<string> requestedLoadSetNames,
            IEnumerable<string> existingLoadSetNames,
            IEnumerable<ShellUniformLoadSetAreaAssignmentDto> assignments,
            Func<ShellUniformLoadSetAreaAssignmentDto, string> resolveAreaObjectName,
            string storyName)
        {
            return CreatePlan(
                requestedLoadSetNames,
                existingLoadSetNames,
                assignments,
                resolveAreaObjectName,
                string.IsNullOrWhiteSpace(storyName) ? null : new[] { storyName });
        }

        public static ShellUniformLoadSetSelectionPlan CreatePlan(
            IEnumerable<string> requestedLoadSetNames,
            IEnumerable<string> existingLoadSetNames,
            IEnumerable<ShellUniformLoadSetAreaAssignmentDto> assignments,
            Func<ShellUniformLoadSetAreaAssignmentDto, string> resolveAreaObjectName,
            IEnumerable<string> storyNames)
        {
            List<string> requestedNames = NormalizeLoadSetNames(requestedLoadSetNames).ToList();
            Dictionary<string, string> existingNames = CreateCanonicalNames(existingLoadSetNames);
            HashSet<string> requestedKeys = new HashSet<string>(
                requestedNames.Select(NormalizeKey),
                StringComparer.OrdinalIgnoreCase);
            List<string> selectedStoryNames = NormalizeStoryNames(storyNames).ToList();
            HashSet<string> selectedStoryKeys = new HashSet<string>(
                selectedStoryNames.Select(NormalizeKey),
                StringComparer.OrdinalIgnoreCase);

            List<string> unknownNames = requestedNames
                .Where(name => !existingNames.ContainsKey(NormalizeKey(name)))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Dictionary<string, string> matchedLoadSets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> selectedAreaSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> selectedAreas = new List<string>();
            List<string> unresolvedAreas = new List<string>();
            int duplicateShellCount = 0;
            int matchingAssignmentCount = 0;

            foreach (ShellUniformLoadSetAreaAssignmentDto assignment in assignments ?? new ShellUniformLoadSetAreaAssignmentDto[0])
            {
                if (assignment == null)
                {
                    continue;
                }

                if (selectedStoryKeys.Count > 0 && !selectedStoryKeys.Contains(NormalizeKey(assignment.Story)))
                {
                    continue;
                }

                string loadSetKey = NormalizeKey(assignment.LoadSetName);
                if (string.IsNullOrWhiteSpace(loadSetKey) || !requestedKeys.Contains(loadSetKey))
                {
                    continue;
                }

                matchingAssignmentCount++;
                string matchedName;
                if (!existingNames.TryGetValue(loadSetKey, out matchedName))
                {
                    matchedName = NormalizeName(assignment.LoadSetName);
                }

                if (!string.IsNullOrWhiteSpace(matchedName) && !matchedLoadSets.ContainsKey(loadSetKey))
                {
                    matchedLoadSets.Add(loadSetKey, matchedName);
                }

                string areaObjectName = NormalizeName(assignment.UniqueName);
                if (string.IsNullOrWhiteSpace(areaObjectName) && resolveAreaObjectName != null)
                {
                    areaObjectName = NormalizeName(resolveAreaObjectName(assignment));
                }

                if (string.IsNullOrWhiteSpace(areaObjectName))
                {
                    unresolvedAreas.Add(FormatAreaReference(assignment));
                    continue;
                }

                if (selectedAreaSet.Add(areaObjectName))
                {
                    selectedAreas.Add(areaObjectName);
                }
                else
                {
                    duplicateShellCount++;
                }
            }

            ShellUniformLoadSetSelectionPlan plan = new ShellUniformLoadSetSelectionPlan
            {
                RequestedLoadSetNames = requestedNames,
                MatchedLoadSetNames = matchedLoadSets.Values.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToList(),
                UnknownLoadSetNames = unknownNames,
                SelectedStoryName = selectedStoryNames.Count == 1 ? selectedStoryNames[0] : string.Empty,
                SelectedStoryNames = selectedStoryNames,
                AreaObjectNames = selectedAreas,
                UnresolvedAreaReferences = unresolvedAreas,
                MatchingAssignmentCount = matchingAssignmentCount,
                DuplicateShellCount = duplicateShellCount,
                UnresolvedAreaCount = unresolvedAreas.Count
            };

            return plan;
        }

        private static Dictionary<string, string> CreateCanonicalNames(IEnumerable<string> loadSetNames)
        {
            Dictionary<string, string> names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawName in loadSetNames ?? new string[0])
            {
                string name = NormalizeName(rawName);
                string key = NormalizeKey(name);
                if (string.IsNullOrWhiteSpace(key) || names.ContainsKey(key))
                {
                    continue;
                }

                names.Add(key, name);
            }

            return names;
        }

        private static IReadOnlyList<string> NormalizeStoryNames(IEnumerable<string> storyNames)
        {
            List<string> names = CreateCanonicalNames(storyNames)
                .Values
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return names;
        }

        private static string FormatAreaReference(ShellUniformLoadSetAreaAssignmentDto assignment)
        {
            string story = NormalizeName(assignment == null ? null : assignment.Story);
            string label = NormalizeName(assignment == null ? null : assignment.Label);
            if (!string.IsNullOrWhiteSpace(story) || !string.IsNullOrWhiteSpace(label))
            {
                return story + "/" + label;
            }

            string loadSet = NormalizeName(assignment == null ? null : assignment.LoadSetName);
            return string.IsNullOrWhiteSpace(loadSet) ? "(unknown area)" : "Load set " + loadSet;
        }

        private static string NormalizeName(string value)
        {
            string name = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
            return name;
        }

        private static string NormalizeKey(string value)
        {
            string key = NormalizeName(value);
            return key;
        }
    }
}

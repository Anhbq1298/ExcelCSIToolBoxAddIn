using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.ShellUniformLoadSetService
{
    internal sealed class ShellUniformLoadSetTableSchema
    {
        public string SetNameFieldKey { get; set; }

        public string LoadPatternFieldKey { get; set; }

        public string LoadValueFieldKey { get; set; }
    }

    internal static class ShellUniformLoadSetTableSchemaResolver
    {
        public static OperationResult<ShellUniformLoadSetTableSchema> Resolve(
            IReadOnlyList<string> fieldKeysIncluded,
            IReadOnlyList<EtabsShellUniformLoadSetTableService.FieldMetadata> allFieldMetadata)
        {
            var included = CreateIncludedFields(fieldKeysIncluded, allFieldMetadata);
            FieldCandidate setName = PickBest(included, new[] { "UniformLoadSetName", "ShellUniformLoadSetName", "LoadSetName", "SetName", "Name" });
            FieldCandidate loadPattern = PickBest(included, new[] { "LoadPattern", "LoadPatternName", "LoadPat", "PatternName", "Pattern" });
            FieldCandidate loadValue = PickBest(included, new[] { "LoadValue", "UniformLoadValue", "UniformLoad", "Value", "Load", "Magnitude" });

            if (!IsResolved(setName) || !IsResolved(loadPattern) || !IsResolved(loadValue) ||
                HasDuplicate(setName, loadPattern, loadValue))
            {
                return OperationResult<ShellUniformLoadSetTableSchema>.Failure(CreateSchemaMismatchMessage(included));
            }

            return OperationResult<ShellUniformLoadSetTableSchema>.Success(new ShellUniformLoadSetTableSchema
            {
                SetNameFieldKey = setName.Field.FieldKey,
                LoadPatternFieldKey = loadPattern.Field.FieldKey,
                LoadValueFieldKey = loadValue.Field.FieldKey
            });
        }

        private static IReadOnlyList<FieldCandidateSource> CreateIncludedFields(
            IReadOnlyList<string> fieldKeysIncluded,
            IReadOnlyList<EtabsShellUniformLoadSetTableService.FieldMetadata> allFieldMetadata)
        {
            var metadataByKey = new Dictionary<string, EtabsShellUniformLoadSetTableService.FieldMetadata>(StringComparer.OrdinalIgnoreCase);
            foreach (EtabsShellUniformLoadSetTableService.FieldMetadata metadata in allFieldMetadata ?? new EtabsShellUniformLoadSetTableService.FieldMetadata[0])
            {
                if (metadata != null && !string.IsNullOrWhiteSpace(metadata.FieldKey) && !metadataByKey.ContainsKey(metadata.FieldKey))
                {
                    metadataByKey.Add(metadata.FieldKey, metadata);
                }
            }

            var fields = new List<FieldCandidateSource>();
            foreach (string fieldKey in fieldKeysIncluded ?? new string[0])
            {
                EtabsShellUniformLoadSetTableService.FieldMetadata metadata;
                if (!metadataByKey.TryGetValue(fieldKey ?? string.Empty, out metadata))
                {
                    metadata = new EtabsShellUniformLoadSetTableService.FieldMetadata { FieldKey = fieldKey ?? string.Empty };
                }

                fields.Add(new FieldCandidateSource { Field = metadata });
            }

            return fields;
        }

        private static FieldCandidate PickBest(IReadOnlyList<FieldCandidateSource> fields, IReadOnlyList<string> aliases)
        {
            var candidates = new List<FieldCandidate>();
            foreach (FieldCandidateSource source in fields)
            {
                int score = Score(source.Field, aliases);
                candidates.Add(new FieldCandidate { Field = source.Field, Score = score });
            }

            int maxScore = candidates.Count == 0 ? 0 : candidates.Max(candidate => candidate.Score);
            if (maxScore <= 0)
            {
                return null;
            }

            List<FieldCandidate> best = candidates.Where(candidate => candidate.Score == maxScore).ToList();
            return best.Count == 1 ? best[0] : null;
        }

        private static int Score(EtabsShellUniformLoadSetTableService.FieldMetadata field, IReadOnlyList<string> aliases)
        {
            string key = NormalizeFieldKey(field == null ? null : field.FieldKey);
            string name = NormalizeFieldKey(field == null ? null : field.FieldName);
            string description = NormalizeFieldKey(field == null ? null : field.Description);
            int best = 0;

            foreach (string alias in aliases)
            {
                string normalizedAlias = NormalizeFieldKey(alias);
                if (key == normalizedAlias)
                {
                    best = Math.Max(best, 100);
                }

                if (name == normalizedAlias)
                {
                    best = Math.Max(best, 90);
                }

                if (key.Contains(normalizedAlias) && normalizedAlias.Length > 4)
                {
                    best = Math.Max(best, 70);
                }

                if (name.Contains(normalizedAlias) && normalizedAlias.Length > 4)
                {
                    best = Math.Max(best, 60);
                }

                if (description.Contains(normalizedAlias) && normalizedAlias.Length > 4)
                {
                    best = Math.Max(best, 40);
                }
            }

            return best;
        }

        private static bool IsResolved(FieldCandidate candidate)
        {
            return candidate != null && candidate.Field != null && !string.IsNullOrWhiteSpace(candidate.Field.FieldKey);
        }

        private static bool HasDuplicate(params FieldCandidate[] candidates)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FieldCandidate candidate in candidates)
            {
                if (candidate == null || candidate.Field == null)
                {
                    return true;
                }

                if (!keys.Add(candidate.Field.FieldKey))
                {
                    return true;
                }
            }

            return false;
        }

        private static string CreateSchemaMismatchMessage(IReadOnlyList<FieldCandidateSource> fields)
        {
            var message = new StringBuilder();
            message.AppendLine("The schema of the ETABS table \"Shell Uniform Load Sets\" is not recognized.");
            message.AppendLine();
            message.AppendLine("Detected fields:");
            foreach (FieldCandidateSource field in fields)
            {
                string fieldKey = field.Field == null ? string.Empty : field.Field.FieldKey;
                string fieldName = field.Field == null ? string.Empty : field.Field.FieldName;
                message.AppendLine("- " + fieldKey + (string.IsNullOrWhiteSpace(fieldName) ? string.Empty : " (" + fieldName + ")"));
            }

            message.AppendLine();
            message.AppendLine("No changes were applied.");
            return message.ToString();
        }

        private static string NormalizeFieldKey(string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey))
            {
                return string.Empty;
            }

            var normalized = new System.Text.StringBuilder(fieldKey.Length);
            foreach (char c in fieldKey)
            {
                if (char.IsLetterOrDigit(c))
                {
                    normalized.Append(char.ToUpperInvariant(c));
                }
            }

            return normalized.ToString();
        }

        private sealed class FieldCandidateSource
        {
            public EtabsShellUniformLoadSetTableService.FieldMetadata Field { get; set; }
        }

        private sealed class FieldCandidate
        {
            public EtabsShellUniformLoadSetTableService.FieldMetadata Field { get; set; }

            public int Score { get; set; }
        }
    }
}

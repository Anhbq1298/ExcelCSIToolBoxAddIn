using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Infrastructure.Etabs;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.ShellUniformLoadSetService
{
    internal static class EtabsShellUniformLoadSetTableService
    {
        private const string TableKey = "Shell Uniform Load Sets";

        public static OperationResult<ShellUniformLoadSetContextDto> GetContext(ETABSv1.cSapModel sapModel)
        {
            if (sapModel == null)
            {
                return OperationResult<ShellUniformLoadSetContextDto>.Failure("ETABS SapModel is not available.");
            }

            try
            {
                string modelPath = sapModel.GetModelFilename(true);
                string modelFileName = string.IsNullOrWhiteSpace(modelPath) ? "Unsaved Model" : Path.GetFileName(modelPath);

                ETABSv1.eForce forceUnits = ETABSv1.eForce.kN;
                ETABSv1.eLength lengthUnits = ETABSv1.eLength.m;
                ETABSv1.eTemperature temperatureUnits = ETABSv1.eTemperature.C;
                string presentUnits = "Units unavailable";
                int unitRet = sapModel.GetPresentUnits_2(ref forceUnits, ref lengthUnits, ref temperatureUnits);
                if (unitRet == 0)
                {
                    presentUnits = EtabsUnitFormatter.FormatDatabaseUnits(forceUnits, lengthUnits, temperatureUnits);
                }

                int numberNames = 0;
                string[] names = null;
                int loadPatternRet = sapModel.LoadPatterns.GetNameList(ref numberNames, ref names);
                if (loadPatternRet != 0)
                {
                    return OperationResult<ShellUniformLoadSetContextDto>.Failure("Failed to read ETABS load patterns (return code " + loadPatternRet.ToString(CultureInfo.InvariantCulture) + ").");
                }

                var dto = new ShellUniformLoadSetContextDto
                {
                    ModelPath = modelPath,
                    ModelFileName = modelFileName,
                    PresentUnitsText = presentUnits,
                    LoadPatternNames = (names ?? new string[0])
                        .Where(name => !string.IsNullOrWhiteSpace(name))
                        .Select(name => name.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                        .ToList()
                };

                return OperationResult<ShellUniformLoadSetContextDto>.Success(dto);
            }
            catch (Exception ex)
            {
                return OperationResult<ShellUniformLoadSetContextDto>.Failure("Failed to initialize Shell Uniform Load Set context: " + ex.Message);
            }
        }

        public static OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>> GetDefinitions(ETABSv1.cSapModel sapModel)
        {
            if (sapModel == null)
            {
                return OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>>.Failure("ETABS SapModel is not available.");
            }

            try
            {
                TableSnapshot snapshot;
                OperationResult readResult = ReadSnapshot(sapModel, out snapshot);
                if (!readResult.IsSuccess)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>>.Failure(readResult.Message);
                }

                IReadOnlyList<FieldMetadata> fieldMetadata = ReadFieldMetadata(sapModel);
                var schemaResult = ShellUniformLoadSetTableSchemaResolver.Resolve(snapshot.FieldKeysIncluded, fieldMetadata);
                if (!schemaResult.IsSuccess)
                {
                    return OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>>.Failure(schemaResult.Message);
                }

                ShellUniformLoadSetTableSchema schema = schemaResult.Data;
                List<Dictionary<string, string>> records = ParseRecords(snapshot);
                var definitionsByName = new Dictionary<string, ShellUniformLoadSetDefinitionDto>(StringComparer.OrdinalIgnoreCase);

                foreach (Dictionary<string, string> record in records)
                {
                    string setName = NormalizeName(ReadRecordValue(record, schema.SetNameFieldKey));
                    string patternName = NormalizeName(ReadRecordValue(record, schema.LoadPatternFieldKey));
                    string valueText = ReadRecordValue(record, schema.LoadValueFieldKey);
                    double value;

                    if (string.IsNullOrWhiteSpace(setName) ||
                        string.IsNullOrWhiteSpace(patternName) ||
                        !TryParseNumber(valueText, out value))
                    {
                        continue;
                    }

                    ShellUniformLoadSetDefinitionDto definition;
                    if (!definitionsByName.TryGetValue(setName, out definition))
                    {
                        definition = new ShellUniformLoadSetDefinitionDto { Name = setName };
                        definitionsByName.Add(setName, definition);
                    }

                    definition.LoadValuesByPattern[patternName] = value;
                }

                IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions = definitionsByName.Values
                    .OrderBy(definition => definition.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>>.Success(definitions);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<ShellUniformLoadSetDefinitionDto>>.Failure("Could not read Shell Uniform Load Sets: " + ex.Message);
            }
        }

        public static OperationResult<ShellUniformLoadSetApplyResultDto> Apply(
            ETABSv1.cSapModel sapModel,
            IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions)
        {
            if (sapModel == null)
            {
                return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("ETABS SapModel is not available.");
            }

            if (definitions == null)
            {
                return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("Shell Uniform Load Set definitions are required.");
            }

            try
            {
                if (sapModel.GetModelIsLocked())
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("The ETABS model is locked. Unlock the model in ETABS before applying Shell Uniform Load Sets.");
                }

                var validation = ValidateDefinitions(definitions);
                if (!validation.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(validation.Message);
                }

                TableSnapshot originalSnapshot;
                OperationResult readResult = ReadSnapshot(sapModel, out originalSnapshot);
                if (!readResult.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(readResult.Message);
                }

                IReadOnlyList<FieldMetadata> fieldMetadata = ReadFieldMetadata(sapModel);
                var schemaResult = ShellUniformLoadSetTableSchemaResolver.Resolve(originalSnapshot.FieldKeysIncluded, fieldMetadata);
                if (!schemaResult.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(schemaResult.Message);
                }

                ShellUniformLoadSetTableSchema schema = schemaResult.Data;
                List<Dictionary<string, string>> existingRecords = ParseRecords(originalSnapshot);

                var submittedNames = new HashSet<string>(
                    definitions.Select(definition => NormalizeName(definition.Name)),
                    StringComparer.OrdinalIgnoreCase);

                var existingNamesBefore = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Dictionary<string, string> record in existingRecords)
                {
                    string existingName = NormalizeName(ReadRecordValue(record, schema.SetNameFieldKey));
                    if (string.IsNullOrWhiteSpace(existingName))
                    {
                        continue;
                    }

                    existingNamesBefore.Add(existingName);
                }

                var finalRecords = new List<Dictionary<string, string>>();
                int loadEntryCount = 0;
                foreach (ShellUniformLoadSetDefinitionDto definition in definitions)
                {
                    string setName = NormalizeName(definition.Name);
                    foreach (KeyValuePair<string, double> pair in definition.LoadValuesByPattern.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        Dictionary<string, string> record = CreateNewDatabaseRecord(originalSnapshot.FieldKeysIncluded, schema, setName, pair.Key, pair.Value);
                        finalRecords.Add(record);
                        loadEntryCount++;
                    }
                }

                OperationResult stageResult = StageAndApply(sapModel, originalSnapshot, finalRecords);
                if (!stageResult.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(stageResult.Message);
                }

                string importLog;
                int warningCount;
                OperationResult applyResult = ApplyEditedTables(sapModel, out warningCount, out importLog);
                if (!applyResult.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(applyResult.Message);
                }

                TableSnapshot readBackSnapshot;
                OperationResult readBackResult = ReadSnapshot(sapModel, out readBackSnapshot);
                if (!readBackResult.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("ETABS accepted the table operation, but the read-back verification could not read the table: " + readBackResult.Message);
                }

                List<Dictionary<string, string>> readBackRecords = ParseRecords(readBackSnapshot);
                OperationResult verification = VerifyReadBack(readBackRecords, schema, definitions);
                if (!verification.IsSuccess)
                {
                    return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure(
                        "ETABS accepted the table operation, but the read-back verification did not match the requested data.\r\n\r\n" +
                        verification.Message +
                        "\r\n\r\nReview the ETABS Import Log and model data.");
                }

                int createdCount = definitions.Count(definition => !existingNamesBefore.Contains(NormalizeName(definition.Name)));
                int updatedCount = definitions.Count - createdCount;
                int deletedCount = existingNamesBefore.Count(name => !submittedNames.Contains(name));
                var result = new ShellUniformLoadSetApplyResultDto
                {
                    CreatedCount = createdCount,
                    UpdatedCount = updatedCount,
                    DeletedCount = deletedCount,
                    LoadEntryCount = loadEntryCount,
                    WarningCount = warningCount,
                    ImportLog = importLog
                };

                return OperationResult<ShellUniformLoadSetApplyResultDto>.Success(result, "Shell Uniform Load Sets updated.");
            }
            catch (Exception ex)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult<ShellUniformLoadSetApplyResultDto>.Failure("Could not update Shell Uniform Load Sets: " + ex.Message);
            }
        }

        private static OperationResult ValidateDefinitions(IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < definitions.Count; i++)
            {
                ShellUniformLoadSetDefinitionDto definition = definitions[i];
                string name = definition == null ? null : NormalizeName(definition.Name);
                if (string.IsNullOrWhiteSpace(name))
                {
                    return OperationResult.Failure("Grid validation failed: UniformLoadSetName is required.");
                }

                if (!names.Add(name))
                {
                    return OperationResult.Failure("Grid validation failed: duplicate UniformLoadSetName '" + name + "'.");
                }

                if (definition.LoadValuesByPattern == null || definition.LoadValuesByPattern.Count == 0)
                {
                    return OperationResult.Failure("Grid validation failed: load set '" + name + "' must contain at least one load value.");
                }
            }

            return OperationResult.Success();
        }

        private static OperationResult ReadSnapshot(ETABSv1.cSapModel sapModel, out TableSnapshot snapshot)
        {
            snapshot = null;
            int tableVersion = 0;
            string[] fieldKeysIncluded = null;
            int numberRecords = 0;
            string[] tableData = null;

            int ret = sapModel.DatabaseTables.GetTableForEditingArray(
                TableKey,
                string.Empty,
                ref tableVersion,
                ref fieldKeysIncluded,
                ref numberRecords,
                ref tableData);

            if (ret != 0)
            {
                return OperationResult.Failure("Could not read ETABS table \"" + TableKey + "\" (return code " + ret.ToString(CultureInfo.InvariantCulture) + ").");
            }

            if (fieldKeysIncluded == null || fieldKeysIncluded.Length == 0)
            {
                return OperationResult.Failure("The ETABS table \"" + TableKey + "\" returned no field keys.");
            }

            tableData = tableData ?? new string[0];
            int expectedLength = numberRecords * fieldKeysIncluded.Length;
            if (tableData.Length != expectedLength)
            {
                return OperationResult.Failure("The ETABS table \"" + TableKey + "\" returned inconsistent data length. Expected " + expectedLength.ToString(CultureInfo.InvariantCulture) + " value(s), received " + tableData.Length.ToString(CultureInfo.InvariantCulture) + ".");
            }

            snapshot = new TableSnapshot
            {
                TableKey = TableKey,
                TableVersion = tableVersion,
                FieldKeysIncluded = fieldKeysIncluded,
                NumberRecords = numberRecords,
                TableData = tableData
            };

            return OperationResult.Success();
        }

        private static IReadOnlyList<FieldMetadata> ReadFieldMetadata(ETABSv1.cSapModel sapModel)
        {
            int tableVersion = 0;
            int numberFields = 0;
            string[] fieldKeys = null;
            string[] fieldNames = null;
            string[] descriptions = null;
            string[] units = null;
            bool[] isImportable = null;

            int ret = sapModel.DatabaseTables.GetAllFieldsInTable(
                TableKey,
                ref tableVersion,
                ref numberFields,
                ref fieldKeys,
                ref fieldNames,
                ref descriptions,
                ref units,
                ref isImportable);

            if (ret != 0 || fieldKeys == null)
            {
                return new List<FieldMetadata>();
            }

            var fields = new List<FieldMetadata>();
            for (int i = 0; i < fieldKeys.Length; i++)
            {
                fields.Add(new FieldMetadata
                {
                    FieldKey = fieldKeys[i] ?? string.Empty,
                    FieldName = GetArrayValue(fieldNames, i),
                    Description = GetArrayValue(descriptions, i),
                    UnitsString = GetArrayValue(units, i),
                    IsImportable = isImportable != null && i < isImportable.Length && isImportable[i]
                });
            }

            return fields;
        }

        private static List<Dictionary<string, string>> ParseRecords(TableSnapshot snapshot)
        {
            var records = new List<Dictionary<string, string>>();
            int fieldCount = snapshot.FieldKeysIncluded.Length;
            for (int recordIndex = 0; recordIndex < snapshot.NumberRecords; recordIndex++)
            {
                var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
                {
                    int valueIndex = recordIndex * fieldCount + fieldIndex;
                    record[snapshot.FieldKeysIncluded[fieldIndex]] = snapshot.TableData[valueIndex] ?? string.Empty;
                }

                records.Add(record);
            }

            return records;
        }

        private static Dictionary<string, string> CreateNewDatabaseRecord(
            IReadOnlyList<string> fieldKeysIncluded,
            ShellUniformLoadSetTableSchema schema,
            string setName,
            string loadPatternName,
            double loadValue)
        {
            var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string fieldKey in fieldKeysIncluded)
            {
                record[fieldKey] = string.Empty;
            }

            record[schema.SetNameFieldKey] = setName;
            record[schema.LoadPatternFieldKey] = loadPatternName;
            record[schema.LoadValueFieldKey] = FormatInvariant(loadValue);
            return record;
        }

        private static OperationResult StageAndApply(
            ETABSv1.cSapModel sapModel,
            TableSnapshot originalSnapshot,
            IReadOnlyList<Dictionary<string, string>> finalRecords)
        {
            TryCancelTableEditing(sapModel);

            string[] fieldKeysIncluded = originalSnapshot.FieldKeysIncluded.ToArray();
            string[] flattenedData = Flatten(finalRecords, fieldKeysIncluded);
            int tableVersion = originalSnapshot.TableVersion;
            int numberRecords = finalRecords.Count;

            int expectedLength = numberRecords * fieldKeysIncluded.Length;
            if (flattenedData.Length != expectedLength)
            {
                return OperationResult.Failure("Internal validation failed while flattening Shell Uniform Load Sets. Expected " + expectedLength.ToString(CultureInfo.InvariantCulture) + " value(s), prepared " + flattenedData.Length.ToString(CultureInfo.InvariantCulture) + ".");
            }

            int ret = sapModel.DatabaseTables.SetTableForEditingArray(
                originalSnapshot.TableKey,
                ref tableVersion,
                ref fieldKeysIncluded,
                numberRecords,
                ref flattenedData);

            if (ret != 0)
            {
                TryCancelTableEditing(sapModel);
                return OperationResult.Failure("Could not stage ETABS table \"" + TableKey + "\" (return code " + ret.ToString(CultureInfo.InvariantCulture) + "). No changes were applied.");
            }

            return OperationResult.Success();
        }

        private static OperationResult ApplyEditedTables(ETABSv1.cSapModel sapModel, out int warningCount, out string importLog)
        {
            bool fillImportLog = true;
            int numFatalErrors = 0;
            int numErrorMessages = 0;
            int numWarningMessages = 0;
            int numInfoMessages = 0;
            importLog = string.Empty;

            int ret = sapModel.DatabaseTables.ApplyEditedTables(
                fillImportLog,
                ref numFatalErrors,
                ref numErrorMessages,
                ref numWarningMessages,
                ref numInfoMessages,
                ref importLog);

            warningCount = numWarningMessages;

            if (ret != 0 || numFatalErrors != 0 || numErrorMessages != 0)
            {
                TryCancelTableEditing(sapModel);
                var message = new StringBuilder();
                message.AppendLine("Could not update Shell Uniform Load Sets.");
                message.AppendLine();
                message.AppendLine("Fatal errors: " + numFatalErrors.ToString(CultureInfo.InvariantCulture));
                message.AppendLine("Errors: " + numErrorMessages.ToString(CultureInfo.InvariantCulture));
                message.AppendLine("Warnings: " + numWarningMessages.ToString(CultureInfo.InvariantCulture));
                message.AppendLine();
                message.AppendLine("ETABS Import Log:");
                message.AppendLine(importLog ?? string.Empty);
                return OperationResult.Failure(message.ToString());
            }

            return OperationResult.Success();
        }

        private static OperationResult VerifyReadBack(
            IReadOnlyList<Dictionary<string, string>> records,
            ShellUniformLoadSetTableSchema schema,
            IReadOnlyList<ShellUniformLoadSetDefinitionDto> definitions)
        {
            var actual = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);
            var namesAfter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var expectedNames = new HashSet<string>(
                definitions.Select(definition => NormalizeName(definition.Name)),
                StringComparer.OrdinalIgnoreCase);
            foreach (Dictionary<string, string> record in records)
            {
                string setName = NormalizeName(ReadRecordValue(record, schema.SetNameFieldKey));
                if (string.IsNullOrWhiteSpace(setName))
                {
                    continue;
                }

                namesAfter.Add(setName);
                string patternName = NormalizeName(ReadRecordValue(record, schema.LoadPatternFieldKey));
                string valueText = ReadRecordValue(record, schema.LoadValueFieldKey);
                double value;
                if (string.IsNullOrWhiteSpace(patternName) || !TryParseNumber(valueText, out value))
                {
                    continue;
                }

                Dictionary<string, double> patternValues;
                if (!actual.TryGetValue(setName, out patternValues))
                {
                    patternValues = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                    actual[setName] = patternValues;
                }

                patternValues[patternName] = value;
            }

            foreach (ShellUniformLoadSetDefinitionDto definition in definitions)
            {
                string setName = NormalizeName(definition.Name);
                Dictionary<string, double> actualValues;
                if (!actual.TryGetValue(setName, out actualValues))
                {
                    return OperationResult.Failure("Load set '" + setName + "' was not found after apply.");
                }

                foreach (KeyValuePair<string, double> expected in definition.LoadValuesByPattern)
                {
                    double actualValue;
                    if (!actualValues.TryGetValue(expected.Key, out actualValue))
                    {
                        return OperationResult.Failure("Expected record was not found: " + setName + " / " + expected.Key + ".");
                    }

                    double tolerance = 1e-9 * Math.Max(1.0, Math.Abs(expected.Value));
                    if (Math.Abs(actualValue - expected.Value) > tolerance)
                    {
                        return OperationResult.Failure("Value mismatch for " + setName + " / " + expected.Key + ". Expected " + FormatInvariant(expected.Value) + ", read back " + FormatInvariant(actualValue) + ".");
                    }
                }

                foreach (string patternName in actualValues.Keys)
                {
                    if (!definition.LoadValuesByPattern.ContainsKey(patternName))
                    {
                        return OperationResult.Failure("Unexpected old record remained for " + setName + " / " + patternName + ".");
                    }
                }
            }

            foreach (string nameAfter in namesAfter)
            {
                if (!expectedNames.Contains(nameAfter))
                {
                    return OperationResult.Failure("Deleted load set '" + nameAfter + "' was still present after apply.");
                }
            }

            return OperationResult.Success();
        }

        private static string[] Flatten(IReadOnlyList<Dictionary<string, string>> records, IReadOnlyList<string> fieldKeysIncluded)
        {
            var data = new List<string>();
            foreach (Dictionary<string, string> record in records)
            {
                foreach (string fieldKey in fieldKeysIncluded)
                {
                    string value;
                    data.Add(record != null && record.TryGetValue(fieldKey, out value) ? value ?? string.Empty : string.Empty);
                }
            }

            return data.ToArray();
        }

        private static void TryCancelTableEditing(ETABSv1.cSapModel sapModel)
        {
            try
            {
                if (sapModel != null && sapModel.DatabaseTables != null)
                {
                    sapModel.DatabaseTables.CancelTableEditing();
                }
            }
            catch
            {
            }
        }

        private static string ReadRecordValue(Dictionary<string, string> record, string fieldKey)
        {
            if (record == null || string.IsNullOrWhiteSpace(fieldKey))
            {
                return string.Empty;
            }

            string value;
            return record.TryGetValue(fieldKey, out value) ? value ?? string.Empty : string.Empty;
        }

        private static string GetArrayValue(string[] values, int index)
        {
            return values != null && index >= 0 && index < values.Length ? values[index] ?? string.Empty : string.Empty;
        }

        private static string NormalizeName(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static bool TryParseNumber(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static string FormatInvariant(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        private sealed class TableSnapshot
        {
            public string TableKey { get; set; }

            public int TableVersion { get; set; }

            public string[] FieldKeysIncluded { get; set; }

            public int NumberRecords { get; set; }

            public string[] TableData { get; set; }
        }

        internal sealed class FieldMetadata
        {
            public string FieldKey { get; set; }

            public string FieldName { get; set; }

            public string Description { get; set; }

            public string UnitsString { get; set; }

            public bool IsImportable { get; set; }
        }
    }
}

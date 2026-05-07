using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.LoadCombinationService
{
    internal delegate int CSISapModelGetCombinationNames<TSapModel>(
        TSapModel sapModel, 
        ref int numberNames, 
        ref string[] names);

    internal delegate int CSISapModelDeleteCombination<TSapModel>(
        TSapModel sapModel,
        string name);

    internal delegate int CSISapModelAddCombination<TSapModel>(
        TSapModel sapModel,
        string name,
        int combinationType);

    internal delegate int CSISapModelGetPatternNames<TSapModel>(
        TSapModel sapModel,
        ref int numberNames,
        ref string[] names);

    internal delegate int CSISapModelSetCombinationCase<TSapModel>(
        TSapModel sapModel,
        string combinationName,
        int caseNameType,
        string caseName,
        double scaleFactor);

    internal delegate int CSISapModelDeleteCombinationCase<TSapModel>(
        TSapModel sapModel,
        string combinationName,
        int caseNameType,
        string caseName);

    internal delegate string CSISapModelGetCombinationType<TSapModel>(
        TSapModel sapModel,
        string name);

    internal delegate int CSISapModelGetCombinationCases<TSapModel>(
        TSapModel sapModel,
        string name,
        ref int numberItems,
        ref string[] caseNames,
        ref int[] caseTypes,
        ref double[] scaleFactors);

    internal static class CSISapModelLoadCombinationService
    {
        internal static OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>> GetLoadCombinations<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetCombinationNames<TSapModel> getCombinationNames,
            CSISapModelGetCombinationType<TSapModel> getCombinationType)
        {
            int numberNames = 0;
            string[] names = null;

            int ret = getCombinationNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>>.Failure("Failed to get load combination names from model.");
                return errorResult;
            }

            var comboList = new List<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>();
            if (names != null)
            {
                foreach (var name in names)
                {
                    string type = getCombinationType(sapModel, name);
                    comboList.Add(new ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO
                    {
                        Name = name,
                        Type = type
                    });
                }
            }

            var successResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadCombinationDTO>>.Success(comboList);
            return successResult;
        }

        public static OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>> GetLoadCombinationDetails<TSapModel>(
            TSapModel sapModel,
            string combinationName,
            CSISapModelGetCombinationCases<TSapModel> getCombinationCases,
            System.Func<TSapModel, string, int, string> resolveTypeName)
        {
            int numberItems = 0;
            string[] caseNames = null;
            int[] caseTypes = null;
            double[] scaleFactors = null;

            int ret = getCombinationCases(sapModel, combinationName, ref numberItems, ref caseNames, ref caseTypes, ref scaleFactors);
            if (ret != 0)
            {
                var failureResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>>.Failure($"Failed to get load combination details for '{combinationName}'.");
                return failureResult;
            }

            var result = new System.Collections.Generic.List<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>();
            for (int i = 0; i < numberItems; i++)
            {
                string typeName = resolveTypeName(sapModel, caseNames[i], caseTypes[i]);
                result.Add(new ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO
                {
                    LoadCaseName = caseNames[i],
                    LoadCaseType = typeName,
                    ScaleFactor = scaleFactors[i]
                });
            }

            var successResult = OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationItemDTO>>.Success(result);
            return successResult;
        }

        internal static OperationResult DeleteLoadCombinations<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<string> combinationNames,
            CSISapModelDeleteCombination<TSapModel> deleteCombination)
        {
            if (combinationNames == null || combinationNames.Count == 0)
            {
                return OperationResult.Failure("No load combinations selected for deletion.");
            }

            int deletedCount = 0;
            int failedCount = 0;

            foreach (var combo in combinationNames)
            {
                int ret = deleteCombination(sapModel, combo);
                if (ret == 0)
                {
                    deletedCount++;
                }
                else
                {
                    failedCount++;
                }
            }

            if (failedCount > 0)
            {
                return OperationResult.Failure($"Deleted {deletedCount} combinations, but failed to delete {failedCount}.");
            }

            return OperationResult.Success($"Successfully deleted {deletedCount} load combinations.");
        }

        internal static OperationResult<IReadOnlyList<string>> GetLoadPatternNames<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetPatternNames<TSapModel> getPatternNames)
        {
            int numberNames = 0;
            string[] names = null;
            int ret = getPatternNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Failed to get load pattern names from model.");
            }

            var result = new List<string>();
            if (names != null)
            {
                foreach (string name in names)
                {
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        result.Add(name);
                    }
                }
            }

            return OperationResult<IReadOnlyList<string>>.Success(result);
        }

        internal static OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto> GetLoadCombinationMatrix<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetPatternNames<TSapModel> getPatternNames,
            CSISapModelGetCombinationNames<TSapModel> getCombinationNames,
            System.Func<TSapModel, string, int> getCombinationType,
            CSISapModelGetCombinationCases<TSapModel> getCombinationCases)
        {
            var patternResult = GetLoadPatternNames(sapModel, getPatternNames);
            if (!patternResult.IsSuccess)
            {
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto>.Failure(patternResult.Message);
            }

            int numberNames = 0;
            string[] names = null;
            int ret = getCombinationNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto>.Failure("Failed to get load combination names from model.");
            }

            var matrix = new ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto();
            matrix.LoadPatternNames.AddRange(patternResult.Data);

            if (names == null)
            {
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto>.Success(matrix);
            }

            var columnNames = new HashSet<string>(matrix.LoadPatternNames, System.StringComparer.OrdinalIgnoreCase);
            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var row = new ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixRowDto
                {
                    LoadCombinationName = name,
                    CombinationType = getCombinationType(sapModel, name)
                };

                int numberItems = 0;
                string[] caseNames = null;
                int[] caseTypes = null;
                double[] scaleFactors = null;
                int caseRet = getCombinationCases(sapModel, name, ref numberItems, ref caseNames, ref caseTypes, ref scaleFactors);
                if (caseRet == 0 && caseNames != null && scaleFactors != null)
                {
                    for (int i = 0; i < numberItems && i < caseNames.Length && i < scaleFactors.Length; i++)
                    {
                        string caseName = caseNames[i];
                        if (string.IsNullOrWhiteSpace(caseName))
                        {
                            continue;
                        }

                        if (!columnNames.Contains(caseName))
                        {
                            matrix.LoadPatternNames.Add(caseName);
                            columnNames.Add(caseName);
                        }

                        row.Factors[caseName] = scaleFactors[i];
                        if (caseTypes != null && i < caseTypes.Length)
                        {
                            row.FactorCaseTypes[caseName] = caseTypes[i];
                        }
                    }
                }

                matrix.Rows.Add(row);
            }

            return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixDto>.Success(matrix);
        }

        internal static OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto> ApplyLoadCombinationMatrix<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationMatrixRowDto> rows,
            CSISapModelGetCombinationNames<TSapModel> getCombinationNames,
            CSISapModelAddCombination<TSapModel> addCombination,
            CSISapModelDeleteCombination<TSapModel> deleteCombination,
            CSISapModelGetCombinationCases<TSapModel> getCombinationCases,
            CSISapModelDeleteCombinationCase<TSapModel> deleteCombinationCase,
            CSISapModelSetCombinationCase<TSapModel> setCombinationCase)
        {
            var result = new ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto();
            if (rows == null || rows.Count == 0)
            {
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto>.Failure("No load combinations were provided.");
            }

            int numberNames = 0;
            string[] names = null;
            int listRet = getCombinationNames(sapModel, ref numberNames, ref names);
            if (listRet != 0)
            {
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto>.Failure("Failed to get existing load combination names from model.");
            }

            var existing = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            if (names != null)
            {
                foreach (string name in names)
                {
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        existing.Add(name);
                    }
                }
            }

            foreach (var row in rows)
            {
                result.ProcessedCount++;
                string comboName = row == null ? null : row.LoadCombinationName;
                if (string.IsNullOrWhiteSpace(comboName))
                {
                    AddFailure(result, comboName, "Validate", "LoadCombinationName is blank.");
                    continue;
                }

                comboName = comboName.Trim();
                int addRet;
                if (existing.Contains(comboName))
                {
                    int deleteRet = deleteCombination(sapModel, comboName);
                    if (deleteRet != 0)
                    {
                        AddFailure(result, comboName, "Delete existing combination", $"return code {deleteRet}");
                        continue;
                    }
                }

                addRet = addCombination(sapModel, comboName, row.CombinationType);
                if (addRet != 0)
                {
                    AddFailure(result, comboName, "Add combination", $"return code {addRet}");
                    continue;
                }

                int clearRet = ClearExistingCases(sapModel, comboName, getCombinationCases, deleteCombinationCase);
                if (clearRet != 0)
                {
                    AddFailure(result, comboName, "Clear existing items", $"return code {clearRet}");
                    continue;
                }

                bool failed = false;
                if (row.Factors != null)
                {
                    foreach (var factor in row.Factors)
                    {
                        if (string.IsNullOrWhiteSpace(factor.Key) || !factor.Value.HasValue || factor.Value.Value == 0d)
                        {
                            continue;
                        }

                        int caseNameType = 0;
                        if (row.FactorCaseTypes != null && row.FactorCaseTypes.TryGetValue(factor.Key, out int savedCaseNameType))
                        {
                            caseNameType = savedCaseNameType;
                        }

                        int setRet = setCombinationCase(sapModel, comboName, caseNameType, factor.Key.Trim(), factor.Value.Value);
                        if (setRet != 0)
                        {
                            AddFailure(result, comboName, $"Add item '{factor.Key}'", $"return code {setRet}");
                            failed = true;
                            break;
                        }
                    }
                }

                if (!failed)
                {
                    result.SuccessCount++;
                }
            }

            result.FailedCount = result.Failures.Count;
            string message = $"Processed {result.ProcessedCount} combination(s): {result.SuccessCount} successful, {result.FailedCount} failed.";
            if (result.FailedCount > 0)
            {
                message += " " + FormatFailures(result.Failures);
                return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto>.Failure(message);
            }

            return OperationResult<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto>.Success(result, message);
        }

        private static int ClearExistingCases<TSapModel>(
            TSapModel sapModel,
            string comboName,
            CSISapModelGetCombinationCases<TSapModel> getCombinationCases,
            CSISapModelDeleteCombinationCase<TSapModel> deleteCombinationCase)
        {
            int numberItems = 0;
            string[] caseNames = null;
            int[] caseTypes = null;
            double[] scaleFactors = null;
            int getRet = getCombinationCases(sapModel, comboName, ref numberItems, ref caseNames, ref caseTypes, ref scaleFactors);
            if (getRet != 0)
            {
                return getRet;
            }

            if (caseNames == null || caseTypes == null)
            {
                return 0;
            }

            for (int i = 0; i < numberItems && i < caseNames.Length && i < caseTypes.Length; i++)
            {
                int deleteRet = deleteCombinationCase(sapModel, comboName, caseTypes[i], caseNames[i]);
                if (deleteRet != 0)
                {
                    return deleteRet;
                }
            }

            return 0;
        }

        private static void AddFailure(
            ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyResultDto result,
            string combinationName,
            string operationName,
            string reason)
        {
            result.Failures.Add(new ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyFailureDto
            {
                LoadCombinationName = string.IsNullOrWhiteSpace(combinationName) ? "(blank)" : combinationName,
                OperationName = operationName,
                Reason = reason
            });
        }

        private static string FormatFailures(IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.LoadCombinationApplyFailureDto> failures)
        {
            var messages = new List<string>();
            foreach (var failure in failures)
            {
                messages.Add($"{failure.LoadCombinationName}: {failure.OperationName} failed ({failure.Reason}).");
            }

            return string.Join(" ", messages);
        }
    }
}


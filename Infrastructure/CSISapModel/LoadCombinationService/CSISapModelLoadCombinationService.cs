using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel.LoadCombinationService
{
    internal delegate int CSISapModelGetCombinationNames<TSapModel>(
        TSapModel sapModel, 
        ref int numberNames, 
        ref string[] names);

    internal delegate int CSISapModelDeleteCombination<TSapModel>(
        TSapModel sapModel,
        string name);

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
        internal static OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>> GetLoadCombinations<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetCombinationNames<TSapModel> getCombinationNames,
            CSISapModelGetCombinationType<TSapModel> getCombinationType)
        {
            int numberNames = 0;
            string[] names = null;

            int ret = getCombinationNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                var errorResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>>.Failure("Failed to get load combination names from model.");
                return errorResult;
            }

            var comboList = new List<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>();
            if (names != null)
            {
                foreach (var name in names)
                {
                    string type = getCombinationType(sapModel, name);
                    comboList.Add(new ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO
                    {
                        Name = name,
                        Type = type
                    });
                }
            }

            var successResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.CSISapModelLoadCombinationDTO>>.Success(comboList);
            return successResult;
        }

        public static OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>> GetLoadCombinationDetails<TSapModel>(
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
                var failureResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>>.Failure($"Failed to get load combination details for '{combinationName}'.");
                return failureResult;
            }

            var result = new System.Collections.Generic.List<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>();
            for (int i = 0; i < numberItems; i++)
            {
                string typeName = resolveTypeName(sapModel, caseNames[i], caseTypes[i]);
                result.Add(new ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO
                {
                    LoadCaseName = caseNames[i],
                    LoadCaseType = typeName,
                    ScaleFactor = scaleFactors[i]
                });
            }

            var successResult = OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Data.DTOs.LoadCombinationItemDTO>>.Success(result);
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
    }
}

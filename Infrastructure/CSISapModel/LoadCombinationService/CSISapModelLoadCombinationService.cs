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

    internal static class CSISapModelLoadCombinationService
    {
        internal static OperationResult<IReadOnlyList<string>> GetLoadCombinations<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetCombinationNames<TSapModel> getCombinationNames)
        {
            int numberNames = 0;
            string[] names = null;

            int ret = getCombinationNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Failed to get load combination names from model.");
            }

            var comboList = new List<string>();
            if (names != null)
            {
                comboList.AddRange(names);
            }

            return OperationResult<IReadOnlyList<string>>.Success(comboList);
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

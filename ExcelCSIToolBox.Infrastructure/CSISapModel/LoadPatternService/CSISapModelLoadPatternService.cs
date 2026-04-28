using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.LoadPatternService
{
    internal delegate int CSISapModelGetPatternNames<TSapModel>(
        TSapModel sapModel,
        ref int numberNames,
        ref string[] names);

    internal delegate int CSISapModelDeletePattern<TSapModel>(
        TSapModel sapModel,
        string name);

    internal delegate string CSISapModelGetPatternType<TSapModel>(
        TSapModel sapModel,
        string name);

    internal static class CSISapModelLoadPatternService
    {
        internal static OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>> GetLoadPatterns<TSapModel>(
            TSapModel sapModel,
            CSISapModelGetPatternNames<TSapModel> getPatternNames,
            CSISapModelGetPatternType<TSapModel> getPatternType)
        {
            int numberNames = 0;
            string[] names = null;

            int ret = getPatternNames(sapModel, ref numberNames, ref names);
            if (ret != 0)
            {
                return OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>>.Failure("Failed to get load pattern names from model.");
            }

            var patternList = new List<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>();
            if (names != null)
            {
                foreach (var name in names)
                {
                    string type = getPatternType(sapModel, name);
                    patternList.Add(new ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO
                    {
                        Name = name,
                        Type = type
                    });
                }
            }

            return OperationResult<IReadOnlyList<ExcelCSIToolBox.Data.DTOs.CSI.CSISapModelLoadPatternDTO>>.Success(patternList);
        }

        internal static OperationResult DeleteLoadPatterns<TSapModel>(
            TSapModel sapModel,
            IReadOnlyList<string> patternNames,
            CSISapModelDeletePattern<TSapModel> deletePattern)
        {
            if (patternNames == null || patternNames.Count == 0)
            {
                return OperationResult.Failure("No load patterns selected for deletion.");
            }

            int deletedCount = 0;
            int failedCount = 0;

            foreach (var pattern in patternNames)
            {
                int ret = deletePattern(sapModel, pattern);
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
                return OperationResult.Failure($"Deleted {deletedCount} patterns, but failed to delete {failedCount}.");
            }

            return OperationResult.Success($"Successfully deleted {deletedCount} load patterns.");
        }
    }
}


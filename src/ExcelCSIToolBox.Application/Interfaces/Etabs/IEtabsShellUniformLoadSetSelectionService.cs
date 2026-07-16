using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsShellUniformLoadSetSelectionService
    {
        OperationResult<IReadOnlyList<string>> GetLoadSetNames();

        OperationResult<IReadOnlyList<string>> GetStoryNames();

        OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(IReadOnlyList<string> loadSetNames);

        OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            string storyName);

        OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            IReadOnlyList<string> storyNames);

        OperationResult<ShellUniformLoadSetSelectionResultDto> SelectShellsByLoadSets(
            IReadOnlyList<string> loadSetNames,
            IReadOnlyList<string> storyNames,
            IProgress<ShellUniformLoadSetSelectionProgressDto> progress);
    }
}

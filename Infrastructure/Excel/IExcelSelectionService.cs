using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public interface IExcelSelectionService
    {
        OperationResult<IReadOnlyList<string>> ReadSingleColumnTextValues();

        OperationResult<IReadOnlyList<ExcelPointCartesianRow>> ReadPointCartesianRows();
    }
}

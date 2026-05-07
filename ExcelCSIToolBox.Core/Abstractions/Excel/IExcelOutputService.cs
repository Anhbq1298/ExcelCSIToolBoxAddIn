using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Tabular;

namespace ExcelCSIToolBox.Core.Abstractions.Excel
{
    public interface IExcelOutputService
    {
        OperationResult WriteDataFrameToActiveCell(DataFrame dataFrame);

        OperationResult WriteValuesToActiveCell(object[,] values, string successMessage = null);
    }
}



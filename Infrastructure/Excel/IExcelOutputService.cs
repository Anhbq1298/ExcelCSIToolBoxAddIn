using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Core.Tabular;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public interface IExcelOutputService
    {
        OperationResult WriteDataFrameToActiveCell(DataFrame dataFrame);
    }
}

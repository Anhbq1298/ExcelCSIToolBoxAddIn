using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public interface IExcelOutputService
    {
        OperationResult WritePointsToActiveCell(IReadOnlyList<EtabsPointData> points);
    }
}

using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Application.Interfaces.Excel
{
    public interface IExcelExportService
    {
        void ExportTable(EtabsTableResult result);
    }
}

using ExcelCSIToolBox.Core.Models.EtabsTables;

namespace ExcelCSIToolBox.Application.Interfaces.Excel
{
    public interface IExcelExportService
    {
        void ExportTable(EtabsTableResult result);
    }
}

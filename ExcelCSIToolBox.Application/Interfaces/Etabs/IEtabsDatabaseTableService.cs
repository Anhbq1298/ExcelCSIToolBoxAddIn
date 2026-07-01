using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs
{
    public interface IEtabsDatabaseTableService
    {
        Task<EtabsTableResult> GetTableAsync(string tableName);
    }
}

using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults
{
    public interface IEtabsAnalysisResultHandler
    {
        bool CanHandle(string key);

        Task ExecuteAsync(AnalysisResultItem item);
    }
}

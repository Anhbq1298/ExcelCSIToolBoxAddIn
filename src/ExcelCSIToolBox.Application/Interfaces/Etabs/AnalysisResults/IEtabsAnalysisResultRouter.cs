using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults
{
    public interface IEtabsAnalysisResultRouter
    {
        Task ExecuteAsync(AnalysisResultItem item);
    }
}

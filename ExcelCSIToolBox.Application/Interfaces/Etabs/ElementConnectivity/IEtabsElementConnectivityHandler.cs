using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity
{
    public interface IEtabsElementConnectivityHandler
    {
        bool CanHandle(string key);

        Task ExecuteAsync(ElementConnectivityItem item);
    }
}

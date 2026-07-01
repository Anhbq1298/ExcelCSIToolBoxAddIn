using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity
{
    public interface IEtabsElementConnectivityRouter
    {
        Task ExecuteAsync(ElementConnectivityItem item);
    }
}

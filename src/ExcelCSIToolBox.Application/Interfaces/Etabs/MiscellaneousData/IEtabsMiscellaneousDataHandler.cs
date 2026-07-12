using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData
{
    public interface IEtabsMiscellaneousDataHandler
    {
        bool CanHandle(string key);

        Task ExecuteAsync(MiscellaneousDataItem item);
    }
}

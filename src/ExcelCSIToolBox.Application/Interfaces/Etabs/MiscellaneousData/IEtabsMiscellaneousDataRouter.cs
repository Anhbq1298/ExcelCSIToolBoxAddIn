using System.Threading.Tasks;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;

namespace ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData
{
    public interface IEtabsMiscellaneousDataRouter
    {
        Task ExecuteAsync(MiscellaneousDataItem item);
    }
}

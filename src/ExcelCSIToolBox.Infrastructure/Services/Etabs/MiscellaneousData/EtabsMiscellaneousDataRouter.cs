using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs.MiscellaneousData;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.MiscellaneousData
{
    public class EtabsMiscellaneousDataRouter : IEtabsMiscellaneousDataRouter
    {
        private readonly List<IEtabsMiscellaneousDataHandler> _handlers;

        public EtabsMiscellaneousDataRouter(IEnumerable<IEtabsMiscellaneousDataHandler> handlers)
        {
            _handlers = handlers == null
                ? new List<IEtabsMiscellaneousDataHandler>()
                : handlers.ToList();
        }

        public async Task ExecuteAsync(MiscellaneousDataItem item)
        {
            if (item == null)
            {
                return;
            }

            IEtabsMiscellaneousDataHandler handler = _handlers.FirstOrDefault(x => x.CanHandle(item.Key));
            if (handler == null)
            {
                throw new InvalidOperationException("No ETABS miscellaneous data handler found for key: " + item.Key);
            }

            await handler.ExecuteAsync(item);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs.ElementConnectivity;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity
{
    public class EtabsElementConnectivityRouter : IEtabsElementConnectivityRouter
    {
        private readonly List<IEtabsElementConnectivityHandler> _handlers;

        public EtabsElementConnectivityRouter(IEnumerable<IEtabsElementConnectivityHandler> handlers)
        {
            _handlers = handlers == null
                ? new List<IEtabsElementConnectivityHandler>()
                : handlers.ToList();
        }

        public async Task ExecuteAsync(ElementConnectivityItem item)
        {
            if (item == null)
            {
                return;
            }

            IEtabsElementConnectivityHandler handler = _handlers.FirstOrDefault(x => x.CanHandle(item.Key));
            if (handler == null)
            {
                throw new InvalidOperationException("No ETABS element connectivity handler found for key: " + item.Key);
            }

            await handler.ExecuteAsync(item);
        }
    }
}

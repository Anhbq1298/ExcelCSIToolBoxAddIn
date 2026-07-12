using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs.AnalysisResults;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults
{
    public class EtabsAnalysisResultRouter : IEtabsAnalysisResultRouter
    {
        private readonly List<IEtabsAnalysisResultHandler> _handlers;

        public EtabsAnalysisResultRouter(IEnumerable<IEtabsAnalysisResultHandler> handlers)
        {
            _handlers = handlers == null
                ? new List<IEtabsAnalysisResultHandler>()
                : handlers.ToList();
        }

        public async Task ExecuteAsync(AnalysisResultItem item)
        {
            if (item == null)
            {
                return;
            }

            IEtabsAnalysisResultHandler handler = _handlers.FirstOrDefault(x => x.CanHandle(item.Key));
            if (handler == null)
            {
                throw new InvalidOperationException("No ETABS analysis result handler found for key: " + item.Key);
            }

            await handler.ExecuteAsync(item);
        }
    }
}

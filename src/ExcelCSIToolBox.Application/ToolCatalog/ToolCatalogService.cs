using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.Features.Selection;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Services;

namespace ExcelCSIToolBox.Application.ToolCatalog
{
    /// <summary>
    /// Dispatches tool catalog operations to Application use cases.
    /// </summary>
    public class ToolCatalogService : IToolCatalogService
    {
        private readonly CsiServiceLocator _serviceLocator;

        public ToolCatalogService(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _serviceLocator = new CsiServiceLocator(etabsService, sap2000Service);
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFrameNames()
        {
            OperationResult<ICSISapModelConnectionService> serviceResult = _serviceLocator.GetActiveService();
            if (!serviceResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(serviceResult.Message);
            }

            return new GetSelectedFrameNamesUseCase(serviceResult.Data).Execute();
        }
    }
}

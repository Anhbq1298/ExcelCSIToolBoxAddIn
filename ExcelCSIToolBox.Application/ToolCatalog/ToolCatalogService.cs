using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.UseCases;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Application.ToolCatalog
{
    /// <summary>
    /// Dispatches tool catalog operations to Application use cases.
    /// </summary>
    public class ToolCatalogService : IToolCatalogService
    {
        private readonly ICSISapModelConnectionService _etabsService;
        private readonly ICSISapModelConnectionService _sap2000Service;

        public ToolCatalogService(
            ICSISapModelConnectionService etabsService,
            ICSISapModelConnectionService sap2000Service)
        {
            _etabsService = etabsService ?? throw new ArgumentNullException(nameof(etabsService));
            _sap2000Service = sap2000Service ?? throw new ArgumentNullException(nameof(sap2000Service));
        }

        public OperationResult<IReadOnlyList<string>> GetSelectedFrameNames()
        {
            OperationResult<ICSISapModelConnectionService> serviceResult = GetActiveService();
            if (!serviceResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(serviceResult.Message);
            }

            return new GetSelectedFrameNamesUseCase(serviceResult.Data).Execute();
        }

        private OperationResult<ICSISapModelConnectionService> GetActiveService()
        {
            OperationResult<CSISapModelConnectionInfoDTO> etabs = _etabsService.GetCurrentConnection();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            OperationResult<CSISapModelConnectionInfoDTO> sap2000 = _sap2000Service.GetCurrentConnection();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            etabs = _etabsService.TryAttachToRunningInstance();
            if (IsConnected(etabs))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_etabsService);
            }

            sap2000 = _sap2000Service.TryAttachToRunningInstance();
            if (IsConnected(sap2000))
            {
                return OperationResult<ICSISapModelConnectionService>.Success(_sap2000Service);
            }

            return OperationResult<ICSISapModelConnectionService>.Failure("No ETABS or SAP2000 model is attached.");
        }

        private static bool IsConnected(OperationResult<CSISapModelConnectionInfoDTO> result)
        {
            return result != null &&
                   result.IsSuccess &&
                   result.Data != null &&
                   result.Data.IsConnected &&
                   result.Data.SapModel != null;
        }
    }
}

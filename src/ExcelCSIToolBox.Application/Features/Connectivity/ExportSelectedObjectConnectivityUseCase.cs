using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Application.Services;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.EtabsTables;

namespace ExcelCSIToolBox.Application.Features.Connectivity
{
    public sealed class ExportSelectedObjectConnectivityUseCase
    {
        private readonly IEtabsDatabaseTableService _tableService;
        private readonly ISelectedObjectIdentityResolver _identityResolver;

        public ExportSelectedObjectConnectivityUseCase(
            IEtabsDatabaseTableService tableService,
            ISelectedObjectIdentityResolver identityResolver)
        {
            _tableService = tableService ?? throw new ArgumentNullException(nameof(tableService));
            _identityResolver = identityResolver ?? throw new ArgumentNullException(nameof(identityResolver));
        }

        public async Task<OperationResult<PreparedTableExport>> ExecuteAsync(ObjectConnectivityRequest request)
        {
            if (request == null)
            {
                return OperationResult<PreparedTableExport>.Failure("Object connectivity export request is required.");
            }

            if (string.IsNullOrWhiteSpace(request.TableName))
            {
                return OperationResult<PreparedTableExport>.Failure("ETABS object connectivity table name is required.");
            }

            OperationResult<IReadOnlyList<CsiObjectIdentity>> identityResult = _identityResolver.ResolveSelectedObjects();
            if (!identityResult.IsSuccess)
            {
                return OperationResult<PreparedTableExport>.Failure(identityResult.Message);
            }

            EtabsTableResult table = await _tableService.GetTableAsync(request.TableName);
            return SelectedObjectTableFilter.Filter(
                table,
                identityResult.Data,
                request.ObjectCategory,
                string.IsNullOrWhiteSpace(request.DisplayName) ? request.TableName : request.DisplayName);
        }
    }
}

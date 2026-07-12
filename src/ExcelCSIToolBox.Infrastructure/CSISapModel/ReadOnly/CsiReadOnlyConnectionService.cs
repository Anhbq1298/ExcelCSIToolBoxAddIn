using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.ReadOnly
{
    /// <summary>
    /// Read-only connection service that delegates to the existing ETABS or SAP2000
    /// connection adapters. This class never exposes the raw SapModel object.
    ///
    /// The service wraps an ICSISapModelConnectionAdapter so the AI tools can call it
    /// safely without touching COM objects directly.
    /// </summary>
    public class CsiReadOnlyConnectionService : ICsiReadOnlyConnectionService
    {
        // Internally held adapters — one per product. Only one is active at a time.
        private readonly ICSISapModelConnectionAdapter<ETABSv1.cSapModel> _etabsAdapter;
        private readonly ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> _sap2000Adapter;

        // Tracks which product is currently active.
        private string _activeProductName = null;

        public CsiReadOnlyConnectionService()
        {
            _etabsAdapter   = CSISapModelConnectionAdapterFactory.CreateEtabs();
            _sap2000Adapter = CSISapModelConnectionAdapterFactory.CreateSap2000();
        }

        public string ProductName
        {
            get { return _activeProductName ?? "None"; }
        }

        /// <summary>Attach to a running ETABS instance. Does not modify the model.</summary>
        public OperationResult<CSISapModelConnectionInfoDTO> AttachToRunningEtabs()
        {
            OperationResult<CSISapModelConnectionInfoDTO> result = _etabsAdapter.TryAttachToRunningInstance();
            if (result.IsSuccess)
            {
                _activeProductName = "ETABS";
            }

            return result;
        }

        /// <summary>Attach to a running SAP2000 instance. Does not modify the model.</summary>
        public OperationResult<CSISapModelConnectionInfoDTO> AttachToRunningSap2000()
        {
            OperationResult<CSISapModelConnectionInfoDTO> result = _sap2000Adapter.TryAttachToRunningInstance();
            if (result.IsSuccess)
            {
                _activeProductName = "SAP2000";
            }

            return result;
        }

        /// <summary>
        /// Return info about the currently active connection.
        /// Tries ETABS first, then SAP2000, using the existing connection if already attached.
        /// </summary>
        public OperationResult<CSISapModelConnectionInfoDTO> GetCurrentModelInfo()
        {
            // Try the most-recently-attached product first.
            if (string.Equals(_activeProductName, "ETABS", StringComparison.OrdinalIgnoreCase))
            {
                OperationResult<CSISapModelConnectionInfoDTO> etabsResult = _etabsAdapter.GetCurrentConnection();
                if (etabsResult.IsSuccess)
                {
                    return etabsResult;
                }
            }

            if (string.Equals(_activeProductName, "SAP2000", StringComparison.OrdinalIgnoreCase))
            {
                OperationResult<CSISapModelConnectionInfoDTO> sap2000Result = _sap2000Adapter.GetCurrentConnection();
                if (sap2000Result.IsSuccess)
                {
                    return sap2000Result;
                }
            }

            // Auto-detect: try ETABS, then SAP2000.
            OperationResult<CSISapModelConnectionInfoDTO> etabsAttach = _etabsAdapter.TryAttachToRunningInstance();
            if (etabsAttach.IsSuccess)
            {
                _activeProductName = "ETABS";
                return etabsAttach;
            }

            OperationResult<CSISapModelConnectionInfoDTO> sap2000Attach = _sap2000Adapter.TryAttachToRunningInstance();
            if (sap2000Attach.IsSuccess)
            {
                _activeProductName = "SAP2000";
                return sap2000Attach;
            }

            return OperationResult<CSISapModelConnectionInfoDTO>.Failure(
                "No running ETABS or SAP2000 instance found. Please open a model and try again.");
        }

        /// <summary>
        /// Return the current units of the active model. Does not modify the model.
        /// </summary>
        public OperationResult<string> GetPresentUnits()
        {
            OperationResult<CSISapModelConnectionInfoDTO> connResult = GetCurrentModelInfo();
            if (!connResult.IsSuccess)
            {
                return OperationResult<string>.Failure(connResult.Message);
            }

            string units = connResult.Data?.ModelCurrentUnit;
            if (string.IsNullOrWhiteSpace(units))
            {
                units = "Units unavailable";
            }

            return OperationResult<string>.Success(units);
        }
    }
}

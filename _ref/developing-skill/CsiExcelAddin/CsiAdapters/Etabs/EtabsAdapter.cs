using System;
using System.Collections.Generic;
using CsiExcelAddin.Models;
using CsiExcelAddin.Services.Interfaces;

// Early binding reference: add ETABSv1.dll to project references before using these types.
// Do NOT use late binding or dynamic for core ETABS workflows.
// Uncomment the line below after adding the reference:
// using ETABSv1;

namespace CsiExcelAddin.CsiAdapters.Etabs
{
    /// <summary>
    /// ETABS-specific implementation of ICsiProductAdapter.
    /// All ETABS API calls are isolated here â€” no ETABS types leak into shared layers.
    ///
    /// IMPORTANT: Verify every API signature against the official ETABSv1 help file
    /// before implementing. Function names and parameter lists must not be assumed.
    /// </summary>
    public class EtabsAdapter : ICsiProductAdapter
    {
        public string ProductName => "ETABS";

        public ICsiApplicationService ApplicationService { get; }
        public ICsiModelService ModelService { get; }

        public EtabsAdapter()
        {
            // Both services share the same internal ETABS connection instance
            var connection = new EtabsConnection();
            ApplicationService = new EtabsApplicationService(connection);
            ModelService = new EtabsModelService(connection);
        }
    }

    // â”€â”€ Internal connection holder â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    // Wraps the raw ETABS API objects so both services share one live reference.

    internal class EtabsConnection
    {
        // TODO: Replace with actual ETABSv1 types after adding the DLL reference.
        // internal cOAPI EtabsObject { get; private set; }
        // internal cSapModel SapModel { get; private set; }

        internal bool IsAttached { get; private set; }

        /// <summary>
        /// Attaches to a running ETABS instance.
        /// API method and parameter list must be verified against ETABSv1 documentation.
        /// </summary>
        internal AttachResult Attach()
        {
            try
            {
                // TODO: Implement using verified ETABSv1 API.
                // Example (verify before use):
                //   EtabsObject = ETABSv1.Helper.GetObject("CSI.ETABS.API.ETABSObject");
                //   SapModel = EtabsObject.SapModel;
                //   int ret = SapModel.InitializeNewModel();

                IsAttached = true;
                return AttachResult.Ok("Attached to ETABS successfully.");
            }
            catch (Exception ex)
            {
                return AttachResult.Fail($"Failed to attach to ETABS: {ex.Message}");
            }
        }

        internal void Detach()
        {
            // Release COM references explicitly to avoid orphaned ETABS processes
            // TODO: SapModel = null; EtabsObject = null;
            IsAttached = false;
        }
    }

    // â”€â”€ Application service â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    internal class EtabsApplicationService : ICsiApplicationService
    {
        private readonly EtabsConnection _connection;

        internal EtabsApplicationService(EtabsConnection connection)
            => _connection = connection;

        public bool IsAttached => _connection.IsAttached;

        public AttachResult Attach() => _connection.Attach();

        public void Detach() => _connection.Detach();
    }

    // â”€â”€ Model service â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    internal class EtabsModelService : ICsiModelService
    {
        private readonly EtabsConnection _connection;

        internal EtabsModelService(EtabsConnection connection)
            => _connection = connection;

        public string GetModelFileName()
        {
            // TODO: Verify API â€” example (do not assume):
            //   return _connection.SapModel.GetModelFilename();
            throw new NotImplementedException("Verify ETABSv1 API signature before implementing.");
        }

        public string GetCurrentUnits()
        {
            // TODO: Verify API â€” units getter differs between ETABS versions.
            throw new NotImplementedException("Verify ETABSv1 API signature before implementing.");
        }

        public IReadOnlyList<FrameSectionDto> GetFrameSections()
        {
            // TODO: Verify API â€” likely SapModel.PropFrame.GetNameList or similar.
            throw new NotImplementedException("Verify ETABSv1 API signature before implementing.");
        }

        public IReadOnlyList<JointDto> GetJoints()
        {
            // TODO: Verify API â€” likely SapModel.PointObj.GetNameList or similar.
            throw new NotImplementedException("Verify ETABSv1 API signature before implementing.");
        }
    }
}


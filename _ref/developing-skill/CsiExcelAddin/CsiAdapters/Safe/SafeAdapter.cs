using System;
using System.Collections.Generic;
using CsiExcelAddin.Models;
using CsiExcelAddin.Services.Interfaces;

// Early binding reference: add SAFEv1.dll (or equivalent) to project references.
// Uncomment after adding the reference:
// using SAFEv1;

namespace CsiExcelAddin.CsiAdapters.Safe
{
    /// <summary>
    /// SAFE-specific implementation of ICsiProductAdapter.
    /// All SAFE API calls are isolated here.
    ///
    /// NOTE: SAFE's API surface differs significantly from ETABS and SAP2000.
    /// Frame section concepts do not apply — replace ICsiModelService members
    /// with SAFE-relevant operations (slabs, punching, etc.) as needed.
    /// Shared interface methods that have no SAFE equivalent should return
    /// an empty collection or a descriptive not-supported message.
    /// </summary>
    public class SafeAdapter : ICsiProductAdapter
    {
        public string ProductName => "SAFE";

        public ICsiApplicationService ApplicationService { get; }
        public ICsiModelService ModelService { get; }

        public SafeAdapter()
        {
            var connection = new SafeConnection();
            ApplicationService = new SafeApplicationService(connection);
            ModelService = new SafeModelService(connection);
        }
    }

    internal class SafeConnection
    {
        // TODO: Replace with actual SAFE API types after adding the DLL reference.
        internal bool IsAttached { get; private set; }

        internal AttachResult Attach()
        {
            try
            {
                // TODO: Implement using verified SAFE API.
                IsAttached = true;
                return AttachResult.Ok("Attached to SAFE successfully.");
            }
            catch (Exception ex)
            {
                return AttachResult.Fail($"Failed to attach to SAFE: {ex.Message}");
            }
        }

        internal void Detach()
        {
            // TODO: Release COM references
            IsAttached = false;
        }
    }

    internal class SafeApplicationService : ICsiApplicationService
    {
        private readonly SafeConnection _connection;
        internal SafeApplicationService(SafeConnection connection) => _connection = connection;

        public bool IsAttached => _connection.IsAttached;
        public AttachResult Attach() => _connection.Attach();
        public void Detach() => _connection.Detach();
    }

    internal class SafeModelService : ICsiModelService
    {
        private readonly SafeConnection _connection;
        internal SafeModelService(SafeConnection connection) => _connection = connection;

        public string GetModelFileName()
            => throw new NotImplementedException("Verify SAFE API signature before implementing.");

        public string GetCurrentUnits()
            => throw new NotImplementedException("Verify SAFE API signature before implementing.");

        public IReadOnlyList<FrameSectionDto> GetFrameSections()
        {
            // SAFE does not use frame sections — return empty list rather than throwing
            return Array.Empty<FrameSectionDto>();
        }

        public IReadOnlyList<JointDto> GetJoints()
            => throw new NotImplementedException("Verify SAFE API signature before implementing.");
    }
}

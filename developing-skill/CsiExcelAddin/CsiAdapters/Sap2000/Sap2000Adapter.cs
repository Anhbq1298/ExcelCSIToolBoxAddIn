using System;
using System.Collections.Generic;
using CsiExcelAddin.Models;
using CsiExcelAddin.Services.Interfaces;

// Early binding reference: add SAP2000v1.dll to project references before using these types.
// Uncomment after adding the reference:
// using SAP2000v1;

namespace CsiExcelAddin.CsiAdapters.Sap2000
{
    /// <summary>
    /// SAP2000-specific implementation of ICsiProductAdapter.
    /// All SAP2000 API calls are isolated here.
    ///
    /// IMPORTANT: SAP2000 API method signatures may differ from ETABS even when
    /// the operation is conceptually the same. Always verify before implementing.
    /// </summary>
    public class Sap2000Adapter : ICsiProductAdapter
    {
        public string ProductName => "SAP2000";

        public ICsiApplicationService ApplicationService { get; }
        public ICsiModelService ModelService { get; }

        public Sap2000Adapter()
        {
            var connection = new Sap2000Connection();
            ApplicationService = new Sap2000ApplicationService(connection);
            ModelService = new Sap2000ModelService(connection);
        }
    }

    internal class Sap2000Connection
    {
        // TODO: Replace with actual SAP2000v1 types after adding the DLL reference.
        // internal cOAPI Sap2000Object { get; private set; }
        // internal cSapModel SapModel { get; private set; }

        internal bool IsAttached { get; private set; }

        internal AttachResult Attach()
        {
            try
            {
                // TODO: Implement using verified SAP2000v1 API.
                IsAttached = true;
                return AttachResult.Ok("Attached to SAP2000 successfully.");
            }
            catch (Exception ex)
            {
                return AttachResult.Fail($"Failed to attach to SAP2000: {ex.Message}");
            }
        }

        internal void Detach()
        {
            // TODO: Release COM references
            IsAttached = false;
        }
    }

    internal class Sap2000ApplicationService : ICsiApplicationService
    {
        private readonly Sap2000Connection _connection;
        internal Sap2000ApplicationService(Sap2000Connection connection) => _connection = connection;

        public bool IsAttached => _connection.IsAttached;
        public AttachResult Attach() => _connection.Attach();
        public void Detach() => _connection.Detach();
    }

    internal class Sap2000ModelService : ICsiModelService
    {
        private readonly Sap2000Connection _connection;
        internal Sap2000ModelService(Sap2000Connection connection) => _connection = connection;

        public string GetModelFileName()
            => throw new NotImplementedException("Verify SAP2000v1 API signature before implementing.");

        public string GetCurrentUnits()
            => throw new NotImplementedException("Verify SAP2000v1 API signature before implementing.");

        public IReadOnlyList<FrameSectionDto> GetFrameSections()
            => throw new NotImplementedException("Verify SAP2000v1 API signature before implementing.");

        public IReadOnlyList<JointDto> GetJoints()
            => throw new NotImplementedException("Verify SAP2000v1 API signature before implementing.");
    }
}

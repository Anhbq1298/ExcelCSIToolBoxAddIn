using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Workflow
{
    public sealed class WorkflowToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("Workflow", "Workflow_CreateTruss", "Add", "Workflow", new[] { "trussType" }, new string[0], new[] { "truss.generate_howe", "truss.generate_pratt" }, WorkflowIntentHints.Truss, WorkflowParameterRules.WorkflowClarification, true),
                SchemaModuleHelpers.Schema("Workflow", "Workflow_CreateShellsFromSelectedFrames", "Add", "Workflow", new string[0], new string[0], new[] { "execute_csi_request" }, new[] { "create shells from selected frames" }, null, true),
                SchemaModuleHelpers.Schema("Workflow", "Workflow_ExportSelectedObjectsToExcel", "Export", "Workflow", new string[0], new string[0], new string[0], new[] { "export selected objects to excel" }, null, false)
            };
        }
    }
}

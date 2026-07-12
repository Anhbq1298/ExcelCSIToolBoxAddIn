using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.Workflow
{
    public sealed class WorkflowToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema(
                    "Workflow",
                    "Workflow_CreateTruss",
                    "Create",
                    "Truss",
                    new[] { "trussType", "span", "startHeight", "endHeight", "numberOfBays" },
                    new[] { "bayLength", "bottomChordElevation", "sectionName", "originX", "originY", "originZ" },
                    new[] { "truss.generate_howe", "truss.generate_pratt", "create truss", "add truss", "generate truss", "create Howe truss", "mono-slope truss", "Howe truss" },
                    WorkflowIntentHints.Truss,
                    WorkflowParameterRules.CreateTrussClarification,
                    true),
                SchemaModuleHelpers.Schema("Workflow", "Workflow_CreateShellsFromSelectedFrames", "Add", "Workflow", new string[0], new string[0], new[] { "execute_csi_request" }, new[] { "create shells from selected frames" }, null, true),
                SchemaModuleHelpers.Schema("Workflow", "Workflow_ExportSelectedObjectsToExcel", "Export", "Workflow", new string[0], new string[0], new string[0], new[] { "export selected objects to excel" }, null, false)
            };
        }
    }
}

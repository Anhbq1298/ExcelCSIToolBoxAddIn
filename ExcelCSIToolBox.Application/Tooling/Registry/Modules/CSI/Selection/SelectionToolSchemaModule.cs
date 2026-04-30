using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Selection
{
    public sealed class SelectionToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("Selection", "Selection_Clear", "Delete", "Selection", new string[0], new string[0], new[] { "selection.clear" }, SelectionIntentHints.Clear, null, true),
                SchemaModuleHelpers.Schema("Selection", "Selection_SelectByName", "Select", "Selection", new[] { "objectType", "objectNames" }, new string[0], new string[0], new[] { "select by name" }, SelectionParameterRules.SelectByNameClarification, true),
                SchemaModuleHelpers.Schema("Selection", "Selection_GetSelectedObjects", "GetInfo", "Selection", new string[0], new string[0], new[] { "CSI.GetSelectedObjects" }, SelectionIntentHints.GetSelectedObjects, null, false),
                SchemaModuleHelpers.Schema("Selection", "Selection_SelectFromExcelRange", "Select", "Selection", new[] { "objectType", "excelRange" }, new string[0], new string[0], new[] { "select from excel" }, "Please provide the object type and Excel range.", true)
            };
        }
    }
}

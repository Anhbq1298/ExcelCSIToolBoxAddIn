using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.LoadCombination
{
    public sealed class LoadCombinationToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("LoadCombination", "LoadCombination_AddLinearAdd", "Add", "LoadCombination", new[] { "loadCombinationName", "loadCaseNames" }, new string[0], new string[0], new[] { "add load combination" }, LoadCombinationParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("LoadCombination", "LoadCombination_Delete", "Delete", "LoadCombination", new[] { "loadCombinationNames" }, new string[0], new[] { "loads.combinations.delete" }, new[] { "delete load combination" }, "Please provide the load combination name(s) to delete.", true),
                SchemaModuleHelpers.Schema("LoadCombination", "LoadCombination_GetList", "GetInfo", "LoadCombination", new string[0], new string[0], new[] { "loads.combinations.get_all" }, LoadCombinationIntentHints.GetList, null, false)
            };
        }
    }
}

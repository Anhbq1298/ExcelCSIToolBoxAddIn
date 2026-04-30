using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.LoadPattern
{
    public sealed class LoadPatternToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("LoadPattern", "LoadPattern_Add", "Add", "LoadPattern", new[] { "loadPatternName", "loadPatternType" }, new string[0], new string[0], new[] { "add load pattern" }, LoadPatternParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("LoadPattern", "LoadPattern_Delete", "Delete", "LoadPattern", new[] { "loadPatternNames" }, new string[0], new[] { "loads.patterns.delete" }, new[] { "delete load pattern" }, "Please provide the load pattern name(s) to delete.", true),
                SchemaModuleHelpers.Schema("LoadPattern", "LoadPattern_GetList", "GetInfo", "LoadPattern", new string[0], new string[0], new[] { "loads.patterns.get_all" }, LoadPatternIntentHints.GetList, null, false)
            };
        }
    }
}

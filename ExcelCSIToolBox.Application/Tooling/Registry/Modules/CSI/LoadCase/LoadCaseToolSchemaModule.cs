using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.LoadCase
{
    public sealed class LoadCaseToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("LoadCase", "LoadCase_AddLinearStatic", "Add", "LoadCase", new[] { "loadCaseName", "loadPatternName" }, new string[0], new string[0], new[] { "add linear static load case" }, LoadCaseParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("LoadCase", "LoadCase_Delete", "Delete", "LoadCase", new[] { "loadCaseNames" }, new string[0], new string[0], new[] { "delete load case" }, "Please provide the load case name(s) to delete.", true),
                SchemaModuleHelpers.Schema("LoadCase", "LoadCase_GetList", "GetInfo", "LoadCase", new string[0], new string[0], new string[0], LoadCaseIntentHints.GetList, null, false)
            };
        }
    }
}

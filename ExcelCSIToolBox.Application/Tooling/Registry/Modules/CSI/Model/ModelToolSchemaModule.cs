using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Model
{
    public sealed class ModelToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("Model", "Model_GetFileName", "GetInfo", "Model", new string[0], new string[0], new[] { "CSI.GetModelInfo" }, ModelIntentHints.Info, null, false),
                SchemaModuleHelpers.Schema("Model", "Model_GetPresentUnits", "GetInfo", "Model", new string[0], new string[0], new[] { "CSI.GetPresentUnits" }, ModelIntentHints.Units, null, false),
                SchemaModuleHelpers.Schema("Model", "Model_RefreshView", "Update", "Model", new string[0], new string[0], new[] { "csi.refresh_view" }, new[] { "refresh view" }, null, true),
                SchemaModuleHelpers.Schema("Model", "Model_Save", "Export", "Model", new[] { "filePath" }, new string[0], new[] { "file.save_model" }, new[] { "save model" }, ModelParameterRules.SaveClarification, true)
            };
        }
    }
}

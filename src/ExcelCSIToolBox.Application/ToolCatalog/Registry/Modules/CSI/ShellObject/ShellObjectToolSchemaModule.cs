using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.ShellObject
{
    public sealed class ShellObjectToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("ShellObject", "ShellObject_AddByPoint", "Add", "ShellObject", new[] { "shellName", "pointNames" }, new[] { "propertyName" }, new[] { "ShellObj_AddByPoint", "shells.add_by_points" }, ShellObjectIntentHints.Add, ShellObjectParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("ShellObject", "ShellObject_AddByCoordinate", "Add", "ShellObject", new[] { "shellName", "coordinates" }, new[] { "propertyName" }, new[] { "shells.add_by_coordinates" }, ShellObjectIntentHints.Add, ShellObjectParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("ShellObject", "ShellObject_GetSelected", "GetInfo", "ShellObject", new string[0], new string[0], new[] { "shells.get_selected" }, ShellObjectIntentHints.GetSelected, null, false)
            };
        }
    }
}

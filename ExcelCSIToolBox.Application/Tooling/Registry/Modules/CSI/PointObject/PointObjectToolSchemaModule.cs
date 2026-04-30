using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.PointObject
{
    public sealed class PointObjectToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema(
                    "PointObject",
                    "PointObject_AddCartesian",
                    "Add",
                    "PointObject",
                    new[] { "pointName", "x", "y", "z" },
                    new string[0],
                    new[] { "PointObj_AddCartesian", "points.add_by_coordinates", "points.add_cartesian" },
                    PointObjectIntentHints.AddCartesian,
                    PointObjectParameterRules.AddCartesianClarification,
                    true),
                SchemaModuleHelpers.Schema("PointObject", "PointObject_Select", "Select", "PointObject", new[] { "pointNames" }, new string[0], new[] { "points.set_selected" }, PointObjectIntentHints.Select, "Please provide the point name(s) to select.", true),
                SchemaModuleHelpers.Schema("PointObject", "PointObject_GetSelected", "GetInfo", "PointObject", new string[0], new string[0], new[] { "points.get_selected" }, PointObjectIntentHints.GetSelected, null, false),
                SchemaModuleHelpers.Schema("PointObject", "PointObject_GetCoordinates", "GetInfo", "PointObject", new[] { "pointName" }, new string[0], new[] { "points.get_coordinates" }, PointObjectIntentHints.GetCoordinates, "Please provide the point name.", false),
                SchemaModuleHelpers.Schema("PointObject", "PointObject_Delete", "Delete", "PointObject", new[] { "pointNames" }, new string[0], new string[0], PointObjectIntentHints.Delete, "Please provide the point name(s) to delete.", true)
            };
        }
    }
}

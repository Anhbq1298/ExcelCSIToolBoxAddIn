using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.FrameObject
{
    public sealed class FrameObjectToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_AddByPoint", "Add", "FrameObject", new[] { "frameName", "pointI", "pointJ" }, new[] { "sectionName" }, new[] { "FrameObj_AddByPoint", "frames.add_by_points" }, FrameObjectIntentHints.Add, FrameObjectParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_AddByCoordinate", "Add", "FrameObject", new[] { "frameName", "xi", "yi", "zi", "xj", "yj", "zj" }, new[] { "sectionName" }, new[] { "FrameObj_AddByCoordinate", "frames.add_by_coordinates", "frames.add_object" }, FrameObjectIntentHints.Add, FrameObjectParameterRules.AddClarification, true),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_AssignSection", "Assign", "FrameObject", new[] { "frameNames", "sectionName" }, new string[0], new[] { "FrameObj_SetSection", "frames.assign_section" }, FrameObjectIntentHints.AssignSection, FrameObjectParameterRules.AssignSectionClarification, true),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_Select", "Select", "FrameObject", new[] { "frameNames" }, new string[0], new[] { "frames.set_selected" }, FrameObjectIntentHints.Select, "Please provide the frame name(s) to select.", true),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_GetSelected", "GetInfo", "FrameObject", new string[0], new string[0], new[] { "frames.get_selected" }, FrameObjectIntentHints.GetSelected, null, false),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_GetEndpoints", "GetInfo", "FrameObject", new[] { "frameName" }, new string[0], new[] { "frames.get_points" }, FrameObjectIntentHints.GetEndpoints, "Please provide the frame name.", false),
                SchemaModuleHelpers.Schema("FrameObject", "FrameObject_Delete", "Delete", "FrameObject", new[] { "frameNames" }, new string[0], new[] { "frames.delete" }, FrameObjectIntentHints.Delete, "Please provide the frame name(s) to delete.", true)
            };
        }
    }
}

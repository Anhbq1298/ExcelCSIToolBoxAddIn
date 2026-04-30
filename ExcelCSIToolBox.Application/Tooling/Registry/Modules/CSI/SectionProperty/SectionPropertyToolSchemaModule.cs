using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI;

namespace ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.SectionProperty
{
    public sealed class SectionPropertyToolSchemaModule : IToolSchemaModule
    {
        public IEnumerable<ToolSchema> GetSchemas()
        {
            return new[]
            {
                SchemaModuleHelpers.Schema("SectionProperty", "SectionProperty_GetList", "GetInfo", "SectionProperty", new string[0], new string[0], new[] { "frames.get_sections" }, SectionPropertyIntentHints.GetList, null, false),
                SchemaModuleHelpers.Schema("SectionProperty", "SectionProperty_GetFrameSectionList", "GetInfo", "FrameObject", new string[0], new string[0], new[] { "CSI.GetSelectedFrameSections", "frames.get_section" }, new[] { "frame sections" }, null, false),
                SchemaModuleHelpers.Schema("SectionProperty", "SectionProperty_AssignToFrame", "Assign", "FrameObject", new[] { "frameNames", "sectionName" }, new string[0], new[] { "FrameObj_SetSection", "FrameObject_AssignSection", "frames.assign_section" }, SectionPropertyIntentHints.AssignToFrame, SectionPropertyParameterRules.AssignToFrameClarification, true)
            };
        }
    }
}

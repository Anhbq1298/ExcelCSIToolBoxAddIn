using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.FrameObject;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.LoadCase;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.LoadCombination;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.LoadPattern;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Model;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.PointObject;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.SectionProperty;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Selection;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.ShellObject;
using ExcelCSIToolBox.Application.Tooling.Registry.Modules.CSI.Workflow;

namespace ExcelCSIToolBox.Application.Tooling.Registry
{
    public static class DefaultToolSchemaRegistryFactory
    {
        public static ToolSchemaRegistry Create()
        {
            return new ToolSchemaRegistryBuilder()
                .AddModule(new PointObjectToolSchemaModule())
                .AddModule(new FrameObjectToolSchemaModule())
                .AddModule(new ShellObjectToolSchemaModule())
                .AddModule(new LoadPatternToolSchemaModule())
                .AddModule(new LoadCaseToolSchemaModule())
                .AddModule(new LoadCombinationToolSchemaModule())
                .AddModule(new SectionPropertyToolSchemaModule())
                .AddModule(new SelectionToolSchemaModule())
                .AddModule(new ModelToolSchemaModule())
                .AddModule(new WorkflowToolSchemaModule())
                .Build();
        }
    }
}

using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.FrameObject;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.LoadCase;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.LoadCombination;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.LoadPattern;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.Model;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.PointObject;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.SectionProperty;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.Selection;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.ShellObject;
using ExcelCSIToolBox.Application.ToolCatalog.Registry.Modules.CSI.Workflow;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry
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

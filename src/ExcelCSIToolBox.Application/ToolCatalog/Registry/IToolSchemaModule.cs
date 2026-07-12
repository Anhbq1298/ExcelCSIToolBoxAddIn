using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry
{
    public interface IToolSchemaModule
    {
        IEnumerable<ToolSchema> GetSchemas();
    }
}

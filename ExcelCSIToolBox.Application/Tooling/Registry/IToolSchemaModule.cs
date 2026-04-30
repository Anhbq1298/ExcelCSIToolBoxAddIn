using System.Collections.Generic;
using ExcelCSIToolBox.Application.Tooling.Contracts;

namespace ExcelCSIToolBox.Application.Tooling.Registry
{
    public interface IToolSchemaModule
    {
        IEnumerable<ToolSchema> GetSchemas();
    }
}

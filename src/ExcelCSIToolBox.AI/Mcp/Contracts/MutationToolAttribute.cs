using System;

namespace ExcelCSIToolBox.AI.Mcp.Contracts
{
    /// <summary>
    /// Marks an MCP tool as a write-path mutation tool requiring explicit confirmation.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, Inherited = true)]
    public sealed class MutationToolAttribute : Attribute
    {
    }
}

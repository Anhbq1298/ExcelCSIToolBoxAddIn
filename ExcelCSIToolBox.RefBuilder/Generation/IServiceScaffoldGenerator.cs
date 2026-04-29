using System.Collections.Generic;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder.Generation
{
    public interface IServiceScaffoldGenerator
    {
        void Generate(string infrastructureRoot, IReadOnlyList<ApiMethodDefinition> methods);
    }
}

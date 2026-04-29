using System.Collections.Generic;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder.Indexing
{
    public interface IApiIndexBuilder
    {
        IReadOnlyList<ApiMethodDefinition> Build(string productName, string dllPath);
    }
}

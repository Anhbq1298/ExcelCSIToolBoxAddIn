using System.Collections.Generic;

namespace ExcelCSIToolBox.RefBuilder.Parsing
{
    public interface IApiDocParser
    {
        IReadOnlyList<ApiMethodDefinition> Parse(string productName, string dllPath);
    }
}

using System.Collections.Generic;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder.Indexing
{
    public sealed class ApiIndexBuilder : IApiIndexBuilder
    {
        private readonly IApiDocParser _parser;

        public ApiIndexBuilder(IApiDocParser parser)
        {
            _parser = parser;
        }

        public IReadOnlyList<ApiMethodDefinition> Build(string productName, string dllPath)
        {
            return _parser.Parse(productName, dllPath);
        }
    }
}

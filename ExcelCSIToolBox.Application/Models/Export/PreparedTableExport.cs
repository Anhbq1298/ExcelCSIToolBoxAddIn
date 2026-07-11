using System.Collections.Generic;

namespace ExcelCSIToolBox.Application.Models.Export
{
    public sealed class PreparedTableExport
    {
        public string TableName { get; set; }

        public IReadOnlyList<string> Headers { get; set; }

        public IReadOnlyList<IReadOnlyList<object>> Rows { get; set; }

        public int RecordCount { get; set; }
    }
}

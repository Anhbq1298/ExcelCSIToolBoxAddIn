using System.Collections.Generic;

namespace ExcelCSIToolBox.Core.Models.AnalysisResults
{
    public class EtabsTableResult
    {
        public EtabsTableResult()
        {
            Headers = new List<string>();
            Rows = new List<List<string>>();
        }

        public string TableName { get; set; }

        public List<string> Headers { get; set; }

        public List<List<string>> Rows { get; set; }
    }
}

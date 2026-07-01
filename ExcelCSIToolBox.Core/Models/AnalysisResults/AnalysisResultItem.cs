namespace ExcelCSIToolBox.Core.Models.AnalysisResults
{
    public class AnalysisResultItem
    {
        public AnalysisResultItem(string title, string key, string category, string etabsTableName)
        {
            Title = title;
            Key = key;
            Category = category;
            EtabsTableName = etabsTableName;
        }

        public string Title { get; set; }

        public string Key { get; set; }

        public string Category { get; set; }

        public string EtabsTableName { get; set; }
    }
}

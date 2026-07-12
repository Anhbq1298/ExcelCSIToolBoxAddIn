namespace ExcelCSIToolBox.Core.Models.ElementConnectivity
{
    public class ElementConnectivityItem
    {
        public ElementConnectivityItem(string title, string key, string groupName, string etabsTableName)
        {
            Title = title;
            Key = key;
            GroupName = groupName;
            EtabsTableName = etabsTableName;
        }

        public string Title { get; private set; }

        public string Key { get; private set; }

        public string GroupName { get; private set; }

        public string EtabsTableName { get; private set; }
    }
}

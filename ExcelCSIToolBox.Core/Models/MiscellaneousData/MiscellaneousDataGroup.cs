using System.Collections.ObjectModel;

namespace ExcelCSIToolBox.Core.Models.MiscellaneousData
{
    public class MiscellaneousDataGroup
    {
        public MiscellaneousDataGroup(string name)
        {
            Name = name;
            Items = new ObservableCollection<MiscellaneousDataItem>();
        }

        public string Name { get; private set; }

        public ObservableCollection<MiscellaneousDataItem> Items { get; private set; }
    }
}

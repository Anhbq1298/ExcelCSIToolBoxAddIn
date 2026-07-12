using System.Collections.ObjectModel;

namespace ExcelCSIToolBox.Core.Models.ElementConnectivity
{
    public class ElementConnectivityGroup
    {
        public ElementConnectivityGroup(string name)
        {
            Name = name;
            Items = new ObservableCollection<ElementConnectivityItem>();
        }

        public string Name { get; private set; }

        public ObservableCollection<ElementConnectivityItem> Items { get; private set; }
    }
}

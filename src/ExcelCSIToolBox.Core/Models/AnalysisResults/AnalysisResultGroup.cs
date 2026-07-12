using System.Collections.ObjectModel;

namespace ExcelCSIToolBox.Core.Models.AnalysisResults
{
    public class AnalysisResultGroup
    {
        public AnalysisResultGroup(string name)
        {
            Name = name;
            Items = new ObservableCollection<AnalysisResultItem>();
        }

        public string Name { get; set; }

        public ObservableCollection<AnalysisResultItem> Items { get; set; }
    }
}

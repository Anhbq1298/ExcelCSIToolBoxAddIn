using System;
using System.Collections.ObjectModel;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.ElementConnectivity
{
    public static class EtabsElementConnectivityRegistry
    {
        public static ObservableCollection<ElementConnectivityGroup> CreateDefaultGroups()
        {
            ObservableCollection<ElementConnectivityGroup> groups = new ObservableCollection<ElementConnectivityGroup>();

            ElementConnectivityGroup connectivity = new ElementConnectivityGroup("Etabs Object Connectivity");
            connectivity.Items.Add(new ElementConnectivityItem("Point Object Connectivity", "POINT_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Point Object Connectivity"));
            connectivity.Items.Add(new ElementConnectivityItem("Beam Object Connectivity", "BEAM_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Beam Object Connectivity"));
            connectivity.Items.Add(new ElementConnectivityItem("Column Object Connectivity", "COLUMN_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Column Object Connectivity"));
            connectivity.Items.Add(new ElementConnectivityItem("Brace Object Connectivity", "BRACE_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Brace Object Connectivity"));
            connectivity.Items.Add(new ElementConnectivityItem("Floor Object Connectivity", "FLOOR_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Floor Object Connectivity"));
            connectivity.Items.Add(new ElementConnectivityItem("Wall Object Connectivity", "WALL_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Wall Object Connectivity"));

            groups.Add(connectivity);
            return groups;
        }

        public static ElementConnectivityGroup CreateGroupForNavigation(string groupName)
        {
            string normalizedGroupName = NormalizeNavigationGroupName(groupName);
            foreach (ElementConnectivityGroup group in CreateDefaultGroups())
            {
                if (string.Equals(group.Name, normalizedGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    return group;
                }
            }

            return new ElementConnectivityGroup(normalizedGroupName);
        }

        public static ObservableCollection<string> GetSupportedKeysForGenericTableExport()
        {
            ObservableCollection<string> keys = new ObservableCollection<string>();
            foreach (ElementConnectivityGroup group in CreateDefaultGroups())
            {
                foreach (ElementConnectivityItem item in group.Items)
                {
                    keys.Add(item.Key);
                }
            }

            return keys;
        }

        private static string NormalizeNavigationGroupName(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
            {
                return "Etabs Object Connectivity";
            }

            switch (groupName)
            {
                case "Point Object Connectivity":
                case "Beam Object Connectivity":
                case "Column Object Connectivity":
                case "Brace Object Connectivity":
                case "Floor Object Connectivity":
                case "Wall Object Connectivity":
                    return "Etabs Object Connectivity";
                default:
                    return groupName;
            }
        }
    }
}

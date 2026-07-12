using System;
using System.Collections.ObjectModel;
using ExcelCSIToolBox.Core.Models.MiscellaneousData;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.DatabaseTables.MiscellaneousData
{
    public static class EtabsMiscellaneousDataRegistry
    {
        public static ObservableCollection<MiscellaneousDataGroup> CreateDefaultGroups()
        {
            ObservableCollection<MiscellaneousDataGroup> groups = new ObservableCollection<MiscellaneousDataGroup>();

            MiscellaneousDataGroup projectInformation = new MiscellaneousDataGroup("Project Information");
            projectInformation.Items.Add(new MiscellaneousDataItem("Project Information", "PROJECT_INFORMATION", "Project Information", "Project Information"));

            MiscellaneousDataGroup materialList = new MiscellaneousDataGroup("Material List");
            materialList.Items.Add(new MiscellaneousDataItem("Material List by Object Type", "MATERIAL_LIST_BY_OBJECT_TYPE", "Material List", "Material List by Object Type"));
            materialList.Items.Add(new MiscellaneousDataItem("Material List by Section Property", "MATERIAL_LIST_BY_SECTION_PROPERTY", "Material List", "Material List by Section Property"));
            materialList.Items.Add(new MiscellaneousDataItem("Material List by Story", "MATERIAL_LIST_BY_STORY", "Material List", "Material List by Story"));

            groups.Add(projectInformation);
            groups.Add(materialList);

            return groups;
        }

        public static MiscellaneousDataGroup CreateGroupForNavigation(string groupName)
        {
            string normalizedGroupName = NormalizeNavigationGroupName(groupName);
            foreach (MiscellaneousDataGroup group in CreateDefaultGroups())
            {
                if (string.Equals(group.Name, normalizedGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    return group;
                }
            }

            return new MiscellaneousDataGroup(normalizedGroupName);
        }

        public static ObservableCollection<string> GetSupportedKeysForGenericTableExport()
        {
            ObservableCollection<string> keys = new ObservableCollection<string>();
            foreach (MiscellaneousDataGroup group in CreateDefaultGroups())
            {
                foreach (MiscellaneousDataItem item in group.Items)
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
                return "Project Information";
            }

            switch (groupName)
            {
                case "Project Information":
                    return "Project Information";
                case "Material List":
                case "Material List by Object Type":
                case "Material List by Section Property":
                case "Material List by Story":
                    return "Material List";
                default:
                    return groupName;
            }
        }
    }
}

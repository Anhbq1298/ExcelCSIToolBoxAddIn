using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Core.Tabular
{
    public static class CsiTableFieldAliasResolver
    {
        private static readonly string[] AnyObjectAliases =
        {
            "Unique Name",
            "UniqueName",
            "Object",
            "Object Name",
            "ObjectName",
            "Obj",
            "ObjName",
            "Name",
            "Label"
        };

        private static readonly string[] PointAliases =
        {
            "Unique Name",
            "UniqueName",
            "Joint",
            "Joint Name",
            "JointName",
            "Point",
            "Point Name",
            "PointName",
            "Label",
            "Label Name",
            "LabelName"
        };

        private static readonly string[] FrameAliases =
        {
            "Unique Name",
            "UniqueName",
            "Frame",
            "Frame Name",
            "FrameName",
            "Column",
            "Column Name",
            "ColumnName",
            "Beam",
            "Beam Name",
            "BeamName",
            "Brace",
            "Brace Name",
            "BraceName",
            "Element",
            "Element Name",
            "ElementName",
            "Label"
        };

        private static readonly string[] AreaAliases =
        {
            "Unique Name",
            "UniqueName",
            "Area",
            "Area Name",
            "AreaName",
            "Shell",
            "Shell Name",
            "ShellName",
            "Floor",
            "Floor Name",
            "FloorName",
            "Wall",
            "Wall Name",
            "WallName",
            "Element",
            "Element Name",
            "ElementName",
            "Label"
        };

        private static readonly string[] PierAliases =
        {
            "Pier",
            "Pier Name",
            "PierName",
            "Label"
        };

        private static readonly string[] SpandrelAliases =
        {
            "Spandrel",
            "Spandrel Name",
            "SpandrelName",
            "Label"
        };

        public static int FindObjectNameColumn(IReadOnlyList<string> fields, string objectType)
        {
            return FindFirstIndex(fields, GetObjectNameAliases(objectType));
        }

        public static int FindFirstIndex(IReadOnlyList<string> fields, params string[] aliases)
        {
            if (fields == null || aliases == null)
            {
                return -1;
            }

            for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                string normalizedField = NormalizeFieldKey(fields[fieldIndex]);
                foreach (string alias in aliases)
                {
                    if (string.Equals(normalizedField, NormalizeFieldKey(alias), StringComparison.OrdinalIgnoreCase))
                    {
                        return fieldIndex;
                    }
                }
            }

            return -1;
        }

        public static string[] GetObjectNameAliases(string objectType)
        {
            string normalizedType = CsiObjectTypes.Normalize(objectType);
            if (string.Equals(normalizedType, CsiObjectTypes.Point, StringComparison.OrdinalIgnoreCase))
            {
                return PointAliases;
            }

            if (string.Equals(normalizedType, CsiObjectTypes.Frame, StringComparison.OrdinalIgnoreCase))
            {
                return FrameAliases;
            }

            if (string.Equals(normalizedType, CsiObjectTypes.Area, StringComparison.OrdinalIgnoreCase))
            {
                return AreaAliases;
            }

            if (string.Equals(normalizedType, CsiObjectTypes.Pier, StringComparison.OrdinalIgnoreCase))
            {
                return PierAliases;
            }

            if (string.Equals(normalizedType, CsiObjectTypes.Spandrel, StringComparison.OrdinalIgnoreCase))
            {
                return SpandrelAliases;
            }

            return AnyObjectAliases;
        }

        public static string NormalizeFieldKey(string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey))
            {
                return string.Empty;
            }

            var chars = new List<char>();
            foreach (char ch in fieldKey)
            {
                if (char.IsLetterOrDigit(ch))
                {
                    chars.Add(char.ToUpperInvariant(ch));
                }
            }

            return chars.Count == 0 ? string.Empty : new string(chars.ToArray());
        }
    }
}

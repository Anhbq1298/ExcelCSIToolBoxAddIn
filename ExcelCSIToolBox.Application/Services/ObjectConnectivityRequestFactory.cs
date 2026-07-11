using System;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.ElementConnectivity;

namespace ExcelCSIToolBox.Application.Services
{
    public static class ObjectConnectivityRequestFactory
    {
        public static ObjectConnectivityRequest Create(ElementConnectivityItem item)
        {
            return new ObjectConnectivityRequest
            {
                TableName = item == null ? string.Empty : item.EtabsTableName,
                DisplayName = item == null ? string.Empty : item.Title,
                ObjectCategory = ResolveObjectCategory(item)
            };
        }

        private static string ResolveObjectCategory(ElementConnectivityItem item)
        {
            string key = item == null ? string.Empty : item.Key;
            string title = item == null ? string.Empty : item.Title;
            string value = string.IsNullOrWhiteSpace(key) ? title : key;

            if (value.IndexOf("POINT", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return CsiObjectTypes.Point;
            }

            if (value.IndexOf("BEAM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("COLUMN", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("BRACE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return CsiObjectTypes.Frame;
            }

            if (value.IndexOf("FLOOR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("WALL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return CsiObjectTypes.Area;
            }

            return CsiObjectTypes.Unknown;
        }
    }
}

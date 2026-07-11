using System;
using System.Collections.Generic;
using System.Linq;
using ExcelCSIToolBox.Application.Models.Export;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Models.EtabsTables;
using ExcelCSIToolBox.Core.Tabular;

namespace ExcelCSIToolBox.Application.Services
{
    public static class SelectedObjectTableFilter
    {
        public static OperationResult<PreparedTableExport> Filter(
            EtabsTableResult table,
            IEnumerable<CsiObjectIdentity> selectedIdentities,
            string objectCategory,
            string displayName)
        {
            string resolvedDisplayName = string.IsNullOrWhiteSpace(displayName)
                ? "ETABS table"
                : displayName.Trim();

            if (table == null)
            {
                return OperationResult<PreparedTableExport>.Failure("No ETABS table data was returned for " + resolvedDisplayName + ".");
            }

            List<CsiObjectIdentity> identities = FilterIdentities(selectedIdentities, objectCategory);
            if (identities.Count == 0)
            {
                return OperationResult<PreparedTableExport>.Failure(
                    "Select one or more " + GetObjectCategoryLabel(objectCategory) + " objects before exporting " + resolvedDisplayName + ".");
            }

            int objectColumnIndex = CsiTableFieldAliasResolver.FindObjectNameColumn(table.Headers, objectCategory);
            if (objectColumnIndex < 0)
            {
                return OperationResult<PreparedTableExport>.Failure(
                    "ETABS " + resolvedDisplayName + " table did not include an object name field required for selected-object filtering.");
            }

            var rows = new List<IReadOnlyList<object>>();
            foreach (List<string> sourceRow in table.Rows ?? new List<List<string>>())
            {
                string rowObjectName = sourceRow != null && objectColumnIndex < sourceRow.Count
                    ? sourceRow[objectColumnIndex]
                    : string.Empty;
                if (!MatchesAnyIdentity(rowObjectName, identities))
                {
                    continue;
                }

                rows.Add(ConvertRow(sourceRow, table.Headers == null ? 0 : table.Headers.Count));
            }

            var headers = table.Headers == null
                ? new List<string>()
                : new List<string>(table.Headers);

            return OperationResult<PreparedTableExport>.Success(new PreparedTableExport
            {
                TableName = table.TableName,
                Headers = headers,
                Rows = rows,
                RecordCount = rows.Count
            });
        }

        private static List<CsiObjectIdentity> FilterIdentities(
            IEnumerable<CsiObjectIdentity> selectedIdentities,
            string objectCategory)
        {
            string normalizedCategory = CsiObjectTypes.Normalize(objectCategory);
            var identities = selectedIdentities == null
                ? new List<CsiObjectIdentity>()
                : selectedIdentities
                    .Where(identity => identity != null)
                    .Where(identity =>
                        string.Equals(normalizedCategory, CsiObjectTypes.Unknown, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(CsiObjectTypes.Normalize(identity.ObjectType), normalizedCategory, StringComparison.OrdinalIgnoreCase))
                    .ToList();

            return identities;
        }

        private static bool MatchesAnyIdentity(string value, IReadOnlyList<CsiObjectIdentity> identities)
        {
            if (string.IsNullOrWhiteSpace(value) || identities == null)
            {
                return false;
            }

            foreach (CsiObjectIdentity identity in identities)
            {
                if (identity != null && identity.Matches(value))
                {
                    return true;
                }
            }

            return false;
        }

        private static IReadOnlyList<object> ConvertRow(IReadOnlyList<string> sourceRow, int headerCount)
        {
            int columnCount = Math.Max(headerCount, sourceRow == null ? 0 : sourceRow.Count);
            var row = new object[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                row[columnIndex] = sourceRow != null && columnIndex < sourceRow.Count
                    ? sourceRow[columnIndex]
                    : string.Empty;
            }

            return row;
        }

        private static string GetObjectCategoryLabel(string objectCategory)
        {
            string normalized = CsiObjectTypes.Normalize(objectCategory);
            if (string.Equals(normalized, CsiObjectTypes.Point, StringComparison.OrdinalIgnoreCase))
            {
                return "joint";
            }

            if (string.Equals(normalized, CsiObjectTypes.Frame, StringComparison.OrdinalIgnoreCase))
            {
                return "frame";
            }

            if (string.Equals(normalized, CsiObjectTypes.Area, StringComparison.OrdinalIgnoreCase))
            {
                return "area";
            }

            if (string.Equals(normalized, CsiObjectTypes.Pier, StringComparison.OrdinalIgnoreCase))
            {
                return "pier";
            }

            if (string.Equals(normalized, CsiObjectTypes.Spandrel, StringComparison.OrdinalIgnoreCase))
            {
                return "spandrel";
            }

            return "ETABS";
        }
    }
}

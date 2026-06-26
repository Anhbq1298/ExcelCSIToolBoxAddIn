using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelCSIToolBox.Infrastructure.Excel;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelCSIToolBoxAddIn.UI.Helpers
{
    internal sealed class PostprocessingWorkbookState
    {
        public string UnitLabel { get; set; }
        public bool AddHeaders { get; set; }
        public bool UsePickedAnchor { get; set; }
        public string AnchorAddress { get; set; }
        public IReadOnlyList<string> LoadCaseNames { get; set; }
        public IReadOnlyList<string> LoadCombinationNames { get; set; }
    }

    internal static class PostprocessingWorkbookStateStore
    {
        private const string PropertyPrefix = "ExcelCSIToolBox.PostProcessing.";
        private const char ListSeparator = '\u001F';

        internal static PostprocessingWorkbookState Load(string toolKey)
        {
            string serialized = ReadCustomProperty(PropertyPrefix + toolKey);
            var values = ParseValues(serialized);
            return new PostprocessingWorkbookState
            {
                UnitLabel = GetValue(values, "Unit"),
                AddHeaders = string.Equals(GetValue(values, "Headers"), "1", StringComparison.Ordinal),
                UsePickedAnchor = string.Equals(GetValue(values, "Pick"), "1", StringComparison.Ordinal),
                AnchorAddress = GetValue(values, "Anchor"),
                LoadCaseNames = SplitList(GetValue(values, "Cases")),
                LoadCombinationNames = SplitList(GetValue(values, "Combos"))
            };
        }

        internal static void Save(string toolKey, PostprocessingWorkbookState state)
        {
            if (state == null)
            {
                return;
            }

            string serialized = string.Join(";", new[]
            {
                CreateValue("Unit", state.UnitLabel),
                CreateValue("Headers", state.AddHeaders ? "1" : "0"),
                CreateValue("Pick", state.UsePickedAnchor ? "1" : "0"),
                CreateValue("Anchor", state.AnchorAddress),
                CreateValue("Cases", JoinList(state.LoadCaseNames)),
                CreateValue("Combos", JoinList(state.LoadCombinationNames))
            });
            WriteCustomProperty(PropertyPrefix + toolKey, serialized);
        }

        internal static ExcelRange TryGetAnchorCell(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return null;
            }

            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                ExcelWorkbook workbook = excelApp == null ? null : excelApp.ActiveWorkbook;
                int separatorIndex = address.LastIndexOf('!');
                if (workbook == null || separatorIndex <= 0 || separatorIndex >= address.Length - 1)
                {
                    return null;
                }

                string sheetName = address.Substring(0, separatorIndex).Trim().Trim('\'');
                string cellAddress = address.Substring(separatorIndex + 1);
                ExcelWorksheet worksheet = workbook.Worksheets[sheetName] as ExcelWorksheet;
                ExcelRange selectedRange = worksheet == null ? null : worksheet.Range[cellAddress] as ExcelRange;
                return selectedRange == null ? null : selectedRange.Cells[1, 1] as ExcelRange;
            }
            catch
            {
                return null;
            }
        }

        private static string ReadCustomProperty(string propertyName)
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                dynamic workbook = excelApp == null ? null : excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    return string.Empty;
                }

                dynamic properties = workbook.CustomDocumentProperties;
                return Convert.ToString(properties[propertyName].Value);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void WriteCustomProperty(string propertyName, string value)
        {
            try
            {
                var excelApp = ExcelApplicationProvider.GetApplication();
                dynamic workbook = excelApp == null ? null : excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    return;
                }

                dynamic properties = workbook.CustomDocumentProperties;
                try
                {
                    properties[propertyName].Value = value;
                }
                catch
                {
                    properties.Add(propertyName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value);
                }
            }
            catch
            {
                // Workbook state is optional and must never block the ETABS workflow.
            }
        }

        private static IDictionary<string, string> ParseValues(string serialized)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string segment in (serialized ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int separatorIndex = segment.IndexOf('=');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                values[segment.Substring(0, separatorIndex)] = Decode(segment.Substring(separatorIndex + 1));
            }

            return values;
        }

        private static string CreateValue(string key, string value)
        {
            return key + "=" + Encode(value ?? string.Empty);
        }

        private static string GetValue(IDictionary<string, string> values, string key)
        {
            string value;
            return values != null && values.TryGetValue(key, out value) ? value : string.Empty;
        }

        private static string JoinList(IReadOnlyList<string> values)
        {
            return values == null ? string.Empty : string.Join(ListSeparator.ToString(), values.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private static IReadOnlyList<string> SplitList(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? new string[0]
                : value.Split(new[] { ListSeparator }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private static string Decode(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value));
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}

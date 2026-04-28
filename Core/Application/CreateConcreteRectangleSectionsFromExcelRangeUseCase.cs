using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Csi;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CreateConcreteRectangleSectionsFromExcelRangeUseCase
    {
        private readonly ICsiConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public CreateConcreteRectangleSectionsFromExcelRangeUseCase(
            ICsiConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadConcreteRectangleSectionRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedCalls = new List<CsiConcreteRectangleSectionInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var sectionName = Normalize(row.SectionName);
                var materialName = Normalize(row.MaterialName);
                var hText = Normalize(row.HText);
                var bText = Normalize(row.BText);

                if (string.IsNullOrWhiteSpace(sectionName) && string.IsNullOrWhiteSpace(materialName) &&
                    string.IsNullOrWhiteSpace(hText) && string.IsNullOrWhiteSpace(bText))
                {
                    continue;
                }

                if (IsHeaderRow(sectionName, materialName, hText, bText))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(sectionName))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: SectionName is blank.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(materialName))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: MaterialName is blank.");
                    continue;
                }

                if (!TryParseDouble(hText, out double h) || !TryParseDouble(bText, out double b))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: h and b must be numeric.");
                    continue;
                }

                if (h <= 0 || b <= 0)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: h and b must all be > 0.");
                    continue;
                }

                orderedCalls.Add(new CsiConcreteRectangleSectionInput
                {
                    SectionName = sectionName,
                    MaterialName = materialName,
                    H = h,
                    B = b
                });
            }

            if (orderedCalls.Count == 0)
            {
                if (failedRowMessages.Count > 0)
                {
                    return OperationResult.Failure($"Excel parsing failed: 0 sections added, {failedRowMessages.Count} row(s) failed. {string.Join(" ", failedRowMessages)}");
                }
                return OperationResult.Failure("Excel parsing failed: no valid rows were found.");
            }

            var addResult = _connectionService.AddConcreteRectangleSections(orderedCalls);
            if (!addResult.IsSuccess)
            {
                return OperationResult.Failure(addResult.Message);
            }

            var message = addResult.Message;
            if (failedRowMessages.Count > 0)
            {
                message += " " + string.Join(" ", failedRowMessages);
            }

            return OperationResult.Success(message);
        }

        private static bool IsHeaderRow(string s1, string s2, string s3, string s4)
        {
            s1 = (s1 ?? "").ToUpper().Replace(" ", "");
            s2 = (s2 ?? "").ToUpper().Replace(" ", "");
            s3 = (s3 ?? "").ToUpper().Replace(" ", "");
            s4 = (s4 ?? "").ToUpper().Replace(" ", "");

            if (s1 == "SECTIONNAME" && s2 == "MATERIAL" && s3 == "H" && s4 == "B")
            {
                return true;
            }
            return false;
        }

        private static string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static bool TryParseDouble(string value, out double result)
        {
            return double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out result)
                || double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);
        }
    }
}

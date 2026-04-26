using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CreateConcreteCircleSectionsFromExcelRangeUseCase
    {
        private readonly ICsiConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public CreateConcreteCircleSectionsFromExcelRangeUseCase(
            ICsiConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadConcreteCircleSectionRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedCalls = new List<EtabsConcreteCircleSectionInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var sectionName = Normalize(row.SectionName);
                var materialName = Normalize(row.MaterialName);
                var dText = Normalize(row.DText);

                if (string.IsNullOrWhiteSpace(sectionName) && string.IsNullOrWhiteSpace(materialName) &&
                    string.IsNullOrWhiteSpace(dText))
                {
                    continue;
                }

                if (IsHeaderRow(sectionName, materialName, dText))
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

                if (!TryParseDouble(dText, out double d))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: d must be numeric.");
                    continue;
                }

                if (d <= 0)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: d must be > 0.");
                    continue;
                }

                orderedCalls.Add(new EtabsConcreteCircleSectionInput
                {
                    SectionName = sectionName,
                    MaterialName = materialName,
                    D = d
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

            var addResult = _connectionService.AddConcreteCircleSections(orderedCalls);
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

        private static bool IsHeaderRow(string s1, string s2, string s3)
        {
            s1 = (s1 ?? "").ToUpper().Replace(" ", "");
            s2 = (s2 ?? "").ToUpper().Replace(" ", "");
            s3 = (s3 ?? "").ToUpper().Replace(" ", "");

            if (s1 == "SECTIONNAME" && s2 == "MATERIAL" && (s3 == "D" || s3 == "DIAMETER"))
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

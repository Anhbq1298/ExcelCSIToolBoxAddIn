using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class CreateSteelAngleSectionsFromExcelRangeUseCase
    {
        private readonly IEtabsConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public CreateSteelAngleSectionsFromExcelRangeUseCase(
            IEtabsConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadSteelAngleSectionRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedCalls = new List<EtabsSteelAngleSectionInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var sectionName = Normalize(row.SectionName);
                var materialName = Normalize(row.MaterialName);
                var hText = Normalize(row.HText);
                var bText = Normalize(row.BText);
                var twText = Normalize(row.TwText);
                var tfText = Normalize(row.TfText);

                if (string.IsNullOrWhiteSpace(sectionName) && string.IsNullOrWhiteSpace(materialName) &&
                    string.IsNullOrWhiteSpace(hText) && string.IsNullOrWhiteSpace(bText) &&
                    string.IsNullOrWhiteSpace(twText) && string.IsNullOrWhiteSpace(tfText))
                {
                    continue;
                }

                if (IsHeaderRow(sectionName, materialName, hText, bText, twText, tfText))
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

                if (!TryParseDouble(hText, out double h) || !TryParseDouble(bText, out double b) ||
                    !TryParseDouble(twText, out double tw) || !TryParseDouble(tfText, out double tf))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: h, b, tw, tf must be numeric.");
                    continue;
                }

                if (h <= 0 || b <= 0 || tw <= 0 || tf <= 0)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: h, b, tw, tf must all be > 0.");
                    continue;
                }

                if (tw >= b)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: tw must be smaller than b.");
                    continue;
                }

                if (2.0 * tf >= h)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: 2*tf must be smaller than h.");
                    continue;
                }

                orderedCalls.Add(new EtabsSteelAngleSectionInput
                {
                    SectionName = sectionName,
                    MaterialName = materialName,
                    H = h,
                    B = b,
                    Tw = tw,
                    Tf = tf
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

            var addResult = _connectionService.AddSteelAngleSections(orderedCalls);
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

        private static bool IsHeaderRow(string s1, string s2, string s3, string s4, string s5, string s6)
        {
            s1 = (s1 ?? "").ToUpper().Replace(" ", "");
            s2 = (s2 ?? "").ToUpper().Replace(" ", "");
            s3 = (s3 ?? "").ToUpper().Replace(" ", "");
            s4 = (s4 ?? "").ToUpper().Replace(" ", "");
            s5 = (s5 ?? "").ToUpper().Replace(" ", "");
            s6 = (s6 ?? "").ToUpper().Replace(" ", "");

            if (s1 == "SECTIONNAME" && s2 == "MATERIAL")
            {
                if ((s3 == "H" || s3 == "DEPTH") && (s4 == "B" || s4 == "WIDTH") &&
                    (s5 == "TW" || s5 == "WEBTHICKNESS") && (s6 == "TF" || s6 == "FLANGETHICKNESS"))
                {
                    return true;
                }
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

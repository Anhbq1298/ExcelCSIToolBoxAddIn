using ExcelCSIToolBox.Data.Models;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Core.Application
{
    public class CreateSteelPipeSectionsFromExcelRangeUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public CreateSteelPipeSectionsFromExcelRangeUseCase(
            ICSISapModelConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadSteelPipeSectionRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedCalls = new List<CSISapModelSteelPipeSectionInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var sectionName = Normalize(row.SectionName);
                var materialName = Normalize(row.MaterialName);
                var odText = Normalize(row.OutsideDiameterText);
                var twText = Normalize(row.WallThicknessText);

                if (string.IsNullOrWhiteSpace(sectionName) && string.IsNullOrWhiteSpace(materialName) &&
                    string.IsNullOrWhiteSpace(odText) && string.IsNullOrWhiteSpace(twText))
                {
                    continue;
                }

                if (IsHeaderRow(sectionName, materialName, odText, twText))
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

                if (!TryParseDouble(odText, out double od) || !TryParseDouble(twText, out double tw))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: OutsideDiameter and WallThickness must be numeric.");
                    continue;
                }

                if (od <= 0 || tw <= 0)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: OutsideDiameter and WallThickness must all be > 0.");
                    continue;
                }

                if (2.0 * tw >= od)
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: 2 * WallThickness must be smaller than OutsideDiameter.");
                    continue;
                }

                orderedCalls.Add(new CSISapModelSteelPipeSectionInput
                {
                    SectionName = sectionName,
                    MaterialName = materialName,
                    OutsideDiameter = od,
                    WallThickness = tw
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

            var addResult = _connectionService.AddSteelPipeSections(orderedCalls);
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

            if (s1 == "SECTIONNAME" && s2 == "MATERIAL")
            {
                if ((s3 == "OUTSIDEDIAMETER" || s3 == "OD" || s3 == "DIAMETER") && 
                    (s4 == "WALLTHICKNESS" || s4 == "THICKNESS" || s4 == "TW" || s4 == "T"))
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



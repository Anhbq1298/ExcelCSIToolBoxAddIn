using ExcelCSIToolBox.Data.Models;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class AddFrameByCoordinatesFromExcelRangeUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public AddFrameByCoordinatesFromExcelRangeUseCase(
            ICSISapModelConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadFrameByCoordRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedFrameCalls = new List<CSISapModelFrameByCoordInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var uniqueName = Normalize(row.UniqueNameText);
                var section = Normalize(row.SectionText);
                var xiText = Normalize(row.XiText);
                var yiText = Normalize(row.YiText);
                var ziText = Normalize(row.ZiText);
                var xjText = Normalize(row.XjText);
                var yjText = Normalize(row.YjText);
                var zjText = Normalize(row.ZjText);

                if (string.IsNullOrWhiteSpace(uniqueName) &&
                    string.IsNullOrWhiteSpace(section) &&
                    string.IsNullOrWhiteSpace(xiText) &&
                    string.IsNullOrWhiteSpace(yiText) &&
                    string.IsNullOrWhiteSpace(ziText) &&
                    string.IsNullOrWhiteSpace(xjText) &&
                    string.IsNullOrWhiteSpace(yjText) &&
                    string.IsNullOrWhiteSpace(zjText))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(section))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Section is required.");
                    continue;
                }

                if (!TryParseDouble(xiText, out double xi))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Xi value '{row.XiText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(yiText, out double yi))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Yi value '{row.YiText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(ziText, out double zi))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Zi value '{row.ZiText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(xjText, out double xj))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Xj value '{row.XjText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(yjText, out double yj))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Yj value '{row.YjText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(zjText, out double zj))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Zj value '{row.ZjText}' is not a valid number.");
                    continue;
                }

                orderedFrameCalls.Add(new CSISapModelFrameByCoordInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    UniqueName = uniqueName,
                    SectionName = section,
                    Xi = xi,
                    Yi = yi,
                    Zi = zi,
                    Xj = xj,
                    Yj = yj,
                    Zj = zj
                });
            }

            if (orderedFrameCalls.Count == 0)
            {
                if (failedRowMessages.Count > 0)
                {
                    return OperationResult.Failure(
                        $"Excel parsing failed: 0 frame(s) added successfully, {failedRowMessages.Count} row(s) failed. {string.Join(" ", failedRowMessages)}");
                }

                return OperationResult.Failure("Excel parsing failed: no valid rows were found in the selected range.");
            }

            var addResult = _connectionService.AddFramesByCoordinates(orderedFrameCalls);
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return OperationResult.Failure(addResult.Message);
            }

            foreach (var failedMessage in addResult.Data.FailedRowMessages)
            {
                failedRowMessages.Add(failedMessage);
            }

            var message = $"{addResult.Data.AddedCount} frame(s) added successfully, {failedRowMessages.Count} row(s) failed.";
            if (failedRowMessages.Count > 0)
            {
                message += " " + string.Join(" ", failedRowMessages);
            }

            return OperationResult.Success(message);
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



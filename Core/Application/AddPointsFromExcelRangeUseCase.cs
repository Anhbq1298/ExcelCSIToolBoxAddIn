using ExcelCSIToolBoxAddIn.Data.Models;
using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class AddPointsFromExcelRangeUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public AddPointsFromExcelRangeUseCase(
            ICSISapModelConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadPointCartesianRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            // Intentionally preserve Excel selection order and duplicates.
            // Each valid row becomes one independent ETABS PointObj.AddCartesian call later in the ETABS service.
            var orderedPointCalls = new List<CSISapModelPointCartesianInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var uniqueName = Normalize(row.UniqueNameText);
                var xText = Normalize(row.XText);
                var yText = Normalize(row.YText);
                var zText = Normalize(row.ZText);

                if (string.IsNullOrWhiteSpace(uniqueName) &&
                    string.IsNullOrWhiteSpace(xText) &&
                    string.IsNullOrWhiteSpace(yText) &&
                    string.IsNullOrWhiteSpace(zText))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(uniqueName))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: UniqueName is required.");
                    continue;
                }

                if (!TryParseDouble(xText, out double x))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: X value '{row.XText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(yText, out double y))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Y value '{row.YText}' is not a valid number.");
                    continue;
                }

                if (!TryParseDouble(zText, out double z))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Z value '{row.ZText}' is not a valid number.");
                    continue;
                }

                orderedPointCalls.Add(new CSISapModelPointCartesianInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    UniqueName = uniqueName,
                    X = x,
                    Y = y,
                    Z = z
                });
            }

            if (orderedPointCalls.Count == 0)
            {
                if (failedRowMessages.Count > 0)
                {
                    return OperationResult.Failure(
                        $"Excel parsing failed: 0 point(s) added successfully, {failedRowMessages.Count} row(s) failed. {string.Join(" ", failedRowMessages)}");
                }

                return OperationResult.Failure("Excel parsing failed: no valid rows were found in the selected range.");
            }

            var addResult = _connectionService.AddPointsByCartesian(orderedPointCalls);
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return OperationResult.Failure(addResult.Message);
            }

            foreach (var failedMessage in addResult.Data.FailedRowMessages)
            {
                failedRowMessages.Add(failedMessage);
            }

            var message = $"{addResult.Data.AddedCount} point(s) added successfully, {failedRowMessages.Count} row(s) failed.";
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


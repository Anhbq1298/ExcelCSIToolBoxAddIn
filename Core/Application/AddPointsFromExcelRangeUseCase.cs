using System.Collections.Generic;
using System.Globalization;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class AddPointsFromExcelRangeUseCase
    {
        private readonly IEtabsConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public AddPointsFromExcelRangeUseCase(
            IEtabsConnectionService connectionService,
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

            var validPoints = new List<EtabsPointCartesianInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var name = Normalize(row.NameText);
                var xText = Normalize(row.XText);
                var yText = Normalize(row.YText);
                var zText = Normalize(row.ZText);

                if (string.IsNullOrWhiteSpace(name) &&
                    string.IsNullOrWhiteSpace(xText) &&
                    string.IsNullOrWhiteSpace(yText) &&
                    string.IsNullOrWhiteSpace(zText))
                {
                    continue;
                }

                if (!TryParseDouble(xText, out double x) ||
                    !TryParseDouble(yText, out double y) ||
                    !TryParseDouble(zText, out double z))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: X, Y, and Z must be valid numbers.");
                    continue;
                }

                validPoints.Add(new EtabsPointCartesianInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    Name = name,
                    X = x,
                    Y = y,
                    Z = z
                });
            }

            if (validPoints.Count == 0)
            {
                if (failedRowMessages.Count > 0)
                {
                    return OperationResult.Failure($"0 point(s) added successfully, {failedRowMessages.Count} row(s) failed. {string.Join(" ", failedRowMessages)}");
                }

                return OperationResult.Failure("No valid rows were found in the selected range.");
            }

            var addResult = _connectionService.AddPointsByCartesian(validPoints);
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

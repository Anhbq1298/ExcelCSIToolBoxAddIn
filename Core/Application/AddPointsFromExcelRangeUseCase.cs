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
            var failedRows = new List<int>();

            foreach (var row in rowResult.Data)
            {
                if (!TryParseDouble(row.XText, out var x) ||
                    !TryParseDouble(row.YText, out var y) ||
                    !TryParseDouble(row.ZText, out var z))
                {
                    failedRows.Add(row.ExcelRowNumber);
                    continue;
                }

                validPoints.Add(new EtabsPointCartesianInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    Name = row.NameText,
                    X = x,
                    Y = y,
                    Z = z
                });
            }

            if (validPoints.Count == 0)
            {
                return OperationResult.Failure("No valid rows were found. Please verify X, Y, Z are numeric.");
            }

            var addResult = _connectionService.AddPointsCartesian(validPoints);
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return OperationResult.Failure(addResult.Message);
            }

            foreach (var failedRow in addResult.Data.FailedRows)
            {
                if (!failedRows.Contains(failedRow))
                {
                    failedRows.Add(failedRow);
                }
            }

            var message = $"{addResult.Data.AddedCount} point(s) added successfully, {failedRows.Count} row(s) failed.";
            if (failedRows.Count > 0)
            {
                message += $" Failed Excel row(s): {string.Join(", ", failedRows)}.";
            }

            return OperationResult.Success(message);
        }

        private static bool TryParseDouble(string text, out double value)
        {
            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out value);
        }
    }
}

using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Infrastructure.Etabs;
using ExcelCSIToolBoxAddIn.Infrastructure.Excel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class AddFramesByPointFromExcelRangeUseCase
    {
        private readonly IEtabsConnectionService _connectionService;
        private readonly IExcelSelectionService _excelSelectionService;

        public AddFramesByPointFromExcelRangeUseCase(
            IEtabsConnectionService connectionService,
            IExcelSelectionService excelSelectionService)
        {
            _connectionService = connectionService;
            _excelSelectionService = excelSelectionService;
        }

        public OperationResult Execute()
        {
            var rowResult = _excelSelectionService.ReadFrameByPointRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var orderedFrameCalls = new List<EtabsFrameByPointInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var uniqueName = Normalize(row.UniqueNameText);
                var section = Normalize(row.SectionText);
                var point1 = Normalize(row.Point1Text);
                var point2 = Normalize(row.Point2Text);

                if (string.IsNullOrWhiteSpace(uniqueName) &&
                    string.IsNullOrWhiteSpace(section) &&
                    string.IsNullOrWhiteSpace(point1) &&
                    string.IsNullOrWhiteSpace(point2))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(section))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Section is required.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(point1))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Point1 is required.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(point2))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: Point2 is required.");
                    continue;
                }

                orderedFrameCalls.Add(new EtabsFrameByPointInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    UniqueName = uniqueName,
                    SectionName = section,
                    Point1Name = point1,
                    Point2Name = point2
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

            var addResult = _connectionService.AddFramesByPoint(orderedFrameCalls);
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
    }
}

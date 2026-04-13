using System.Collections.Generic;
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
            var rowResult = _excelSelectionService.ReadFrameByPointRows();
            if (!rowResult.IsSuccess)
            {
                return OperationResult.Failure(rowResult.Message);
            }

            var validFrames = new List<EtabsFrameByPointInput>();
            var failedRowMessages = new List<string>();

            foreach (var row in rowResult.Data)
            {
                var uniqueName = Normalize(row.UniqueNameText);
                var section = Normalize(row.SectionText);
                var pointIName = Normalize(row.PointINameText);
                var pointJName = Normalize(row.PointJNameText);

                if (string.IsNullOrWhiteSpace(uniqueName) &&
                    string.IsNullOrWhiteSpace(section) &&
                    string.IsNullOrWhiteSpace(pointIName) &&
                    string.IsNullOrWhiteSpace(pointJName))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(pointIName) || string.IsNullOrWhiteSpace(pointJName))
                {
                    failedRowMessages.Add($"Row {row.ExcelRowNumber}: PointIName and PointJName are required.");
                    continue;
                }

                validFrames.Add(new EtabsFrameByPointInput
                {
                    ExcelRowNumber = row.ExcelRowNumber,
                    UniqueName = uniqueName,
                    Section = section,
                    PointIName = pointIName,
                    PointJName = pointJName
                });
            }

            if (validFrames.Count == 0)
            {
                if (failedRowMessages.Count > 0)
                {
                    return OperationResult.Failure($"0 object(s) added successfully, {failedRowMessages.Count} row(s) failed. {string.Join(" ", failedRowMessages)}");
                }

                return OperationResult.Failure("No valid rows were found in the selected range.");
            }

            var addResult = _connectionService.AddFramesByPointPairs(validFrames);
            if (!addResult.IsSuccess || addResult.Data == null)
            {
                return OperationResult.Failure(addResult.Message);
            }

            foreach (var failedMessage in addResult.Data.FailedRowMessages)
            {
                failedRowMessages.Add(failedMessage);
            }

            var message = $"{addResult.Data.AddedCount} object(s) added successfully, {failedRowMessages.Count} row(s) failed.";
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

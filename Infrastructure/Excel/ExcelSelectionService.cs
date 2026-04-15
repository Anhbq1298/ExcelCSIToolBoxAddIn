using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using Microsoft.Office.Interop.Excel;

namespace ExcelCSIToolBoxAddIn.Infrastructure.Excel
{
    public class ExcelSelectionService : IExcelSelectionService
    {
        public OperationResult<IReadOnlyList<string>> ReadSingleColumnTextValues()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<string>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            if (selection.Columns.Count != 1)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("Please select exactly 1 column and N rows.");
            }

            var values = new List<string>();
            int rowCount = selection.Rows.Count;
            object rawValues = selection.Value2;

            for (int row = 1; row <= rowCount; row++)
            {
                string value = ReadCellText(rawValues, selection, row, 1);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    values.Add(value);
                }
            }

            if (values.Count == 0)
            {
                return OperationResult<IReadOnlyList<string>>.Failure("The selected Excel range does not contain any non-empty values.");
            }

            return OperationResult<IReadOnlyList<string>>.Success(values);
        }

        public OperationResult<IReadOnlyList<ExcelPointCartesianRow>> ReadPointCartesianRows()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<ExcelPointCartesianRow>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            int rowCount = selection.Rows.Count;
            int columnCount = selection.Columns.Count;

            if (rowCount < 1)
            {
                return OperationResult<IReadOnlyList<ExcelPointCartesianRow>>.Failure("Excel range validation failed: please select at least 1 row.");
            }

            if (columnCount != 4)
            {
                return OperationResult<IReadOnlyList<ExcelPointCartesianRow>>.Failure(
                    $"Excel range validation failed: expected exactly 4 columns (UniqueName, X, Y, Z), but found {columnCount}.");
            }

            var rows = new List<ExcelPointCartesianRow>();
            object rawValues = selection.Value2;

            for (int row = 1; row <= rowCount; row++)
            {
                rows.Add(new ExcelPointCartesianRow
                {
                    ExcelRowNumber = selection.Row + row - 1,
                    UniqueNameText = ReadCellText(rawValues, selection, row, 1),
                    XText = ReadCellText(rawValues, selection, row, 2),
                    YText = ReadCellText(rawValues, selection, row, 3),
                    ZText = ReadCellText(rawValues, selection, row, 4)
                });
            }

            if (rows.Count == 0)
            {
                return OperationResult<IReadOnlyList<ExcelPointCartesianRow>>.Failure("Excel range validation failed: please select at least one row.");
            }

            return OperationResult<IReadOnlyList<ExcelPointCartesianRow>>.Success(rows);
        }

        public OperationResult<IReadOnlyList<ExcelFrameByCoordRow>> ReadFrameByCoordRows()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByCoordRow>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            int rowCount = selection.Rows.Count;
            int columnCount = selection.Columns.Count;

            if (rowCount < 1)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByCoordRow>>.Failure("Excel range validation failed: please select at least 1 row.");
            }

            if (columnCount < 8)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByCoordRow>>.Failure(
                    $"Excel range validation failed: expected at least 8 columns (UniqueName, Section, Xi, Yi, Zi, Xj, Yj, Zj), but found {columnCount}.");
            }

            var rows = new List<ExcelFrameByCoordRow>();
            object rawValues = selection.Value2;

            for (int row = 1; row <= rowCount; row++)
            {
                rows.Add(new ExcelFrameByCoordRow
                {
                    ExcelRowNumber = selection.Row + row - 1,
                    UniqueNameText = ReadCellText(rawValues, selection, row, 1),
                    SectionText = ReadCellText(rawValues, selection, row, 2),
                    XiText = ReadCellText(rawValues, selection, row, 3),
                    YiText = ReadCellText(rawValues, selection, row, 4),
                    ZiText = ReadCellText(rawValues, selection, row, 5),
                    XjText = ReadCellText(rawValues, selection, row, 6),
                    YjText = ReadCellText(rawValues, selection, row, 7),
                    ZjText = ReadCellText(rawValues, selection, row, 8)
                });
            }

            if (rows.Count == 0)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByCoordRow>>.Failure("Excel range validation failed: please select at least one row.");
            }

            return OperationResult<IReadOnlyList<ExcelFrameByCoordRow>>.Success(rows);
        }

        public OperationResult<IReadOnlyList<ExcelFrameByPointRow>> ReadFrameByPointRows()
        {
            var selectionResult = GetActiveSelection();
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            int rowCount = selection.Rows.Count;
            int columnCount = selection.Columns.Count;

            if (rowCount < 1)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure("Excel range validation failed: please select at least 1 row.");
            }

            if (columnCount < 4)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure(
                    $"Excel range validation failed: expected at least 4 columns (UniqueName, Section, Point1, Point2), but found {columnCount}.");
            }

            var rows = new List<ExcelFrameByPointRow>();
            object rawValues = selection.Value2;

            for (int row = 1; row <= rowCount; row++)
            {
                rows.Add(new ExcelFrameByPointRow
                {
                    ExcelRowNumber = selection.Row + row - 1,
                    UniqueNameText = ReadCellText(rawValues, selection, row, 1),
                    SectionText = ReadCellText(rawValues, selection, row, 2),
                    Point1Text = ReadCellText(rawValues, selection, row, 3),
                    Point2Text = ReadCellText(rawValues, selection, row, 4)
                });
            }

            if (rows.Count == 0)
            {
                return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Failure("Excel range validation failed: please select at least one row.");
            }

            return OperationResult<IReadOnlyList<ExcelFrameByPointRow>>.Success(rows);
        }

        private static OperationResult<Range> GetActiveSelection()
        {
            try
            {
                Application excelApp = Globals.ExcelCSIToolBoxAddin.Application;
                if (excelApp == null)
                {
                    return OperationResult<Range>.Failure("Excel application is not available.");
                }

                var selection = excelApp.Selection as Range;
                if (selection == null)
                {
                    return OperationResult<Range>.Failure("Please select a range in Excel and try again.");
                }

                return OperationResult<Range>.Success(selection);
            }
            catch (Exception)
            {
                return OperationResult<Range>.Failure("Unable to read the current Excel selection.");
            }
        }

        private static string ReadCellText(object rawValues, Range selection, int row, int column)
        {
            try
            {
                if (rawValues is object[,])
                {
                    object[,] matrix = (object[,])rawValues;
                    return ToTrimmedText(matrix[row, column]);
                }

                if (rawValues != null && row == 1 && column == 1)
                {
                    return ToTrimmedText(rawValues);
                }

                var cell = selection.Cells[row, column] as Range;
                if (cell == null)
                {
                    return null;
                }

                return ToTrimmedText(cell.Value2);
            }
            catch
            {
                return null;
            }
        }

        private static string ToTrimmedText(object value)
        {
            if (value == null)
            {
                return null;
            }

            return Convert.ToString(value)?.Trim();
        }
    }
}

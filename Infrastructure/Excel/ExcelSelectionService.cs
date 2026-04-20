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
            var selectionResult = GetActiveSelection("Select a single column range:", "Select Items");
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
            var selectionResult = GetActiveSelection("Select a 4-column range:\r\nUniqueName | X | Y | Z", "Select Point Cartesian Range");
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
            var selectionResult = GetActiveSelection("Select an 8-column range:\r\nUniqueName | Section | Xi | Yi | Zi | Xj | Yj | Zj", "Select Frame by Coordinates Range");
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
            var selectionResult = GetActiveSelection("Select a 4-column range:\r\nUniqueName | Section | Point1 | Point2", "Select Frame by Points Range");
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

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelISectionRow>> ReadSteelISectionRows()
        {
            return ReadRows(6, "SectionName, Material, h, b, tw, tf", "Select a 6-column range:\r\nSectionName | Material | h | b | tw | tf", "Select Steel I-Section Input Range", (rawValues, selection, row) => new ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelISectionRow
            {
                ExcelRowNumber = selection.Row + row - 1,
                SectionName = ReadCellText(rawValues, selection, row, 1),
                MaterialName = ReadCellText(rawValues, selection, row, 2),
                HText = ReadCellText(rawValues, selection, row, 3),
                BText = ReadCellText(rawValues, selection, row, 4),
                TwText = ReadCellText(rawValues, selection, row, 5),
                TfText = ReadCellText(rawValues, selection, row, 6)
            });
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelChannelSectionRow>> ReadSteelChannelSectionRows()
        {
            return ReadRows(6, "SectionName, Material, h, b, tw, tf", "Select a 6-column range:\r\nSectionName | Material | h | b | tw | tf", "Select Steel Channel Section Input Range", (rawValues, selection, row) => new ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelChannelSectionRow
            {
                ExcelRowNumber = selection.Row + row - 1,
                SectionName = ReadCellText(rawValues, selection, row, 1),
                MaterialName = ReadCellText(rawValues, selection, row, 2),
                HText = ReadCellText(rawValues, selection, row, 3),
                BText = ReadCellText(rawValues, selection, row, 4),
                TwText = ReadCellText(rawValues, selection, row, 5),
                TfText = ReadCellText(rawValues, selection, row, 6)
            });
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelAngleSectionRow>> ReadSteelAngleSectionRows()
        {
            return ReadRows(6, "SectionName, Material, h, b, tw, tf", "Select a 6-column range:\r\nSectionName | Material | h | b | tw | tf", "Select Steel Angle Section Input Range", (rawValues, selection, row) => new ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelAngleSectionRow
            {
                ExcelRowNumber = selection.Row + row - 1,
                SectionName = ReadCellText(rawValues, selection, row, 1),
                MaterialName = ReadCellText(rawValues, selection, row, 2),
                HText = ReadCellText(rawValues, selection, row, 3),
                BText = ReadCellText(rawValues, selection, row, 4),
                TwText = ReadCellText(rawValues, selection, row, 5),
                TfText = ReadCellText(rawValues, selection, row, 6)
            });
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelPipeSectionRow>> ReadSteelPipeSectionRows()
        {
            return ReadRows(4, "SectionName, Material, OutsideDiameter, WallThickness", "Select a 4-column range:\r\nSectionName | Material | OutsideDiameter | WallThickness", "Select Steel Pipe Section Input Range", (rawValues, selection, row) => new ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelPipeSectionRow
            {
                ExcelRowNumber = selection.Row + row - 1,
                SectionName = ReadCellText(rawValues, selection, row, 1),
                MaterialName = ReadCellText(rawValues, selection, row, 2),
                OutsideDiameterText = ReadCellText(rawValues, selection, row, 3),
                WallThicknessText = ReadCellText(rawValues, selection, row, 4)
            });
        }

        public OperationResult<IReadOnlyList<ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelTubeSectionRow>> ReadSteelTubeSectionRows()
        {
            return ReadRows(5, "SectionName, Material, h, b, t", "Select a 5-column range:\r\nSectionName | Material | h | b | t", "Select Steel Tube Section Input Range", (rawValues, selection, row) => new ExcelCSIToolBoxAddIn.Core.Tabular.ExcelSteelTubeSectionRow
            {
                ExcelRowNumber = selection.Row + row - 1,
                SectionName = ReadCellText(rawValues, selection, row, 1),
                MaterialName = ReadCellText(rawValues, selection, row, 2),
                HText = ReadCellText(rawValues, selection, row, 3),
                BText = ReadCellText(rawValues, selection, row, 4),
                TText = ReadCellText(rawValues, selection, row, 5)
            });
        }

        private OperationResult<IReadOnlyList<T>> ReadRows<T>(int expectedColumns, string expectedColumnsDesc, string prompt, string title, System.Func<object, Range, int, T> rowMapper)
        {
            var selectionResult = GetActiveSelection(prompt, title);
            if (!selectionResult.IsSuccess)
            {
                return OperationResult<IReadOnlyList<T>>.Failure(selectionResult.Message);
            }

            var selection = selectionResult.Data;
            int rowCount = selection.Rows.Count;
            int columnCount = selection.Columns.Count;

            if (rowCount < 1)
            {
                return OperationResult<IReadOnlyList<T>>.Failure("Excel range validation failed: please select at least 1 row.");
            }

            if (columnCount < expectedColumns)
            {
                return OperationResult<IReadOnlyList<T>>.Failure(
                    $"Excel range validation failed: expected at least {expectedColumns} columns ({expectedColumnsDesc}), but found {columnCount}.");
            }

            var rows = new List<T>();
            object rawValues = selection.Value2;

            for (int row = 1; row <= rowCount; row++)
            {
                rows.Add(rowMapper(rawValues, selection, row));
            }

            if (rows.Count == 0)
            {
                return OperationResult<IReadOnlyList<T>>.Failure("Excel range validation failed: please select at least one valid row.");
            }

            return OperationResult<IReadOnlyList<T>>.Success(rows);
        }

        private static OperationResult<Range> GetActiveSelection(string prompt, string title)
        {
            try
            {
                Application excelApp = Globals.ExcelCSIToolBoxAddin.Application;
                if (excelApp == null)
                {
                    return OperationResult<Range>.Failure("Excel application is not available.");
                }

                object result = excelApp.InputBox(prompt, title, Type: 8);
                if (result is bool b && !b)
                {
                    return OperationResult<Range>.Failure("Action canceled by user.");
                }

                var selection = result as Range;
                if (selection == null)
                {
                    return OperationResult<Range>.Failure("Please select a valid range in Excel and try again.");
                }

                return OperationResult<Range>.Success(selection);
            }
            catch (Exception)
            {
                return OperationResult<Range>.Failure("Unable to read the current Excel selection or action canceled.");
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

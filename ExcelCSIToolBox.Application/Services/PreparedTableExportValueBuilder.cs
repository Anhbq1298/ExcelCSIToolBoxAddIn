using System.Collections.Generic;
using ExcelCSIToolBox.Application.Models.Export;

namespace ExcelCSIToolBox.Application.Services
{
    public static class PreparedTableExportValueBuilder
    {
        public static object[,] BuildValues(PreparedTableExport preparedExport, bool addHeaders)
        {
            int columnCount = GetColumnCount(preparedExport);
            int dataRowCount = preparedExport == null || preparedExport.Rows == null ? 0 : preparedExport.Rows.Count;
            int rowCount = dataRowCount + (addHeaders ? 1 : 0);
            if (columnCount == 0 || rowCount == 0)
            {
                return new object[0, 0];
            }

            var values = new object[rowCount, columnCount];
            int dataRowOffset = addHeaders ? 1 : 0;
            if (addHeaders && preparedExport != null && preparedExport.Headers != null)
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[0, columnIndex] = columnIndex < preparedExport.Headers.Count
                        ? preparedExport.Headers[columnIndex]
                        : string.Empty;
                }
            }

            if (preparedExport == null || preparedExport.Rows == null)
            {
                return values;
            }

            for (int rowIndex = 0; rowIndex < preparedExport.Rows.Count; rowIndex++)
            {
                IReadOnlyList<object> row = preparedExport.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[rowIndex + dataRowOffset, columnIndex] =
                        row != null && columnIndex < row.Count ? row[columnIndex] : string.Empty;
                }
            }

            return values;
        }

        private static int GetColumnCount(PreparedTableExport preparedExport)
        {
            if (preparedExport == null)
            {
                return 0;
            }

            int columnCount = preparedExport.Headers == null ? 0 : preparedExport.Headers.Count;
            if (preparedExport.Rows != null)
            {
                foreach (IReadOnlyList<object> row in preparedExport.Rows)
                {
                    if (row != null && row.Count > columnCount)
                    {
                        columnCount = row.Count;
                    }
                }
            }

            return columnCount;
        }
    }
}

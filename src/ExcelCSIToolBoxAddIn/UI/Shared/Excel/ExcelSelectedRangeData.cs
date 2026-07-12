namespace ExcelCSIToolBoxAddIn.UI.Forms
{
    internal sealed class ExcelSelectedRangeData
    {
        private readonly object[,] _values;

        public ExcelSelectedRangeData(object[,] values, int rowCount, int columnCount)
        {
            _values = values;
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        public int RowCount { get; private set; }

        public int ColumnCount { get; private set; }

        public object GetValue(int row, int column)
        {
            return _values[row, column];
        }
    }
}

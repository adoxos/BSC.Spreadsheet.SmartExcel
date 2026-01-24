using System;

namespace SmartExcel
{
    public sealed class ExcelRangeData
    {
        public ExcelRangeData(object[,] values, int rowCount, int columnCount)
        {
            Values = values ?? throw new ArgumentNullException(nameof(values));
            RowCount = rowCount;
            ColumnCount = columnCount;
        }

        public object[,] Values { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }
    }
}

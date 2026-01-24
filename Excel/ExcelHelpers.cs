using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SmartExcel
{
    /// <summary>
    /// Utility methods shared across Excel add-in forms and operations.
    /// </summary>
    public static class ExcelHelpers
    {
        /// <summary>
        /// Converts a 1-based column index to an Excel column letter (A, B, ..., AA).
        /// </summary>
        public static string GetColumnLetter(int columnNumber)
        {
            int dividend = columnNumber;
            var columnName = new System.Text.StringBuilder();

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName.Insert(0, Convert.ToChar(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString();
        }

        /// <summary>
        /// Returns true when a worksheet has only the default empty cell.
        /// </summary>
        public static bool IsUsedRangeEmpty(dynamic usedRange)
        {
            if (usedRange == null)
            {
                return true;
            }

            if (usedRange.Rows.Count == 1 && usedRange.Columns.Count == 1)
            {
                var value = usedRange.Value2;
                return value == null || string.IsNullOrWhiteSpace(value.ToString());
            }

            return false;
        }

        /// <summary>
        /// Loads column letters into a combo box, showing a friendly message when there is no data.
        /// </summary>
        public static bool TryPopulateColumnComboBox(dynamic excelApp, ComboBox columnComboBox)
        {
            if (columnComboBox == null)
            {
                throw new ArgumentNullException(nameof(columnComboBox));
            }

            try
            {
                dynamic activeSheet = excelApp.ActiveSheet;
                var usedRange = activeSheet.UsedRange;

                if (IsUsedRangeEmpty(usedRange))
                {
                    MessageBox.Show("No data found in the active sheet.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                int columnCount = usedRange.Columns.Count;
                var items = new List<string>(columnCount);

                for (int i = 1; i <= columnCount; i++)
                {
                    items.Add($"Column {GetColumnLetter(i)}");
                }

                columnComboBox.Items.Clear();
                columnComboBox.Items.AddRange(items.ToArray());

                if (columnComboBox.Items.Count > 0)
                {
                    columnComboBox.SelectedIndex = 0;
                }

                return columnComboBox.Items.Count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading columns: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Normalizes Excel's Value2 return type to a 1-based 2D array with explicit row/column metadata.
        /// </summary>
        internal static ExcelRangeData NormalizeRangeValues(dynamic usedRange)
        {
            object value = usedRange.Value2;

            if (value is object[,] array)
            {
                return new ExcelRangeData(array, usedRange.Rows.Count, usedRange.Columns.Count);
            }

            int rows = usedRange.Rows.Count;
            int columns = usedRange.Columns.Count;

            // Allocate with an extra slot so indices 1..N remain valid, matching Excel's 1-based arrays.
            var singleCellData = new object[rows + 1, columns + 1];
            singleCellData[1, 1] = value;

            return new ExcelRangeData(singleCellData, rows, columns);
        }
    }
}

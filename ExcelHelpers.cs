using System;

namespace SmartExcel
{
    /// <summary>
    /// Utility class for common helper methods
    /// </summary>
    public static class ExcelHelpers
    {
        /// <summary>
        /// Converts a column number to Excel column letter(s)
        /// </summary>
        /// <param name="columnNumber">1-based column number</param>
        /// <returns>Column letter (e.g., A, B, Z, AA, AB)</returns>
        public static string GetColumnLetter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}

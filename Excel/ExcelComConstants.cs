namespace SmartExcel
{
    /// <summary>
    /// Excel COM enumeration values used without referencing the official interop assembly.
    /// </summary>
    internal static class ExcelComConstants
    {
        public const int CalculationAutomatic = -4105; // xlCalculationAutomatic
        public const int CalculationManual = -4135;    // xlCalculationManual
        public const int ShiftDown = -4121;            // xlShiftDown
        public const int FormatFromLeftOrAbove = -4122; // xlFormatFromLeftOrAbove
    }
}

namespace ExcelWriter.OpenXml.Excel
{
    /// <summary>
    /// Excel specific constants
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Maximum number of characters in an Excel worksheet name
        /// </summary>
        public const int SheetNameMaxLength = 31;

        /// <summary>
        /// Number of 'English Metric Units' per cm
        /// </summary>
        public const double EmusPerCentimeter = 360000;

        /// <summary>
        /// Number of 'English Metric Units' per inch
        /// </summary>
        public const double EmusPerInch = 914400;

        /// <summary>
        /// Number of Points per Inch (Don't confuse this with Pixels per Inch)
        /// </summary>
        public const double PointsPerInch = 72;

        /// <summary>
        /// Number of 'English Metric Units' per pixel at 72dpi (or Points per Inch).
        /// </summary>
        public const double EmusPerPoint = 914400 / 72;

        ///// <summary>
        ///// Number of 'English Metric Units' per Excel column unit (almost).
        ///// </summary>
        //public const double EmusPerExcelWidthUnit = 70400;
    }
}

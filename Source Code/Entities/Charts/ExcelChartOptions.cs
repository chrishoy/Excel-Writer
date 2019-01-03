namespace ExcelWriter
{
    /// <summary>
    /// Represents the options that are associated with exporting of charts
    /// </summary>
    public sealed class ExcelChartOptions
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelChartOptions" /> class.
        /// </summary>
        public ExcelChartOptions()
        {
            this.ChartType = ExportChartType.None;
        }

        /// <summary>
        /// Gets or sets a value indicating the name of the chart in excel.<br/>
        /// If not supplied the 1st chart will be used
        /// </summary>
        public string ChartName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the type of chart, required when dealing with dynamic data
        /// </summary>
        public ExportChartType ChartType { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the colours of all series in the chart will be set using the GAM Colour Pallette.
        /// Any other colours specified will be overriden
        /// </summary>
        public bool UseGAMColoursForSeries { get; set; }
    }
}

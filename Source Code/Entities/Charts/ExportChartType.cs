namespace ExcelWriter
{
    /// <summary>
    /// Types of charts that are suppported (or not supported if None).
    /// </summary>
    public enum ExportChartType
    {
        /// <summary>
        /// The chart is not supported
        /// </summary>
        None,

        /// <summary>
        /// The chart is a Bar chart
        /// </summary>
        BarChart,

        /// <summary>
        /// The chart is a Line chart
        /// </summary>
        LineChart,

        /// <summary>
        /// The chart is a Scatter chart
        /// </summary>
        ScatterChart        
    }
}

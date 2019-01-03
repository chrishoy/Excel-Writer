namespace ExcelWriter
{
    using System.Collections.Generic;

    /// <summary>
    /// When implemented on a <see cref="DataPart"/> row or <see cref="TableColumn"/>,
    /// this property determines information about the row/column which<br/>
    /// relates to how it is presented in charts.
    /// </summary>
    public interface IChartMetadata
    {
        /// <summary>
        /// Gets a list of options that will be used when representing this data within a chart.
        /// </summary>
        ChartOptions ChartOptions { get; }
    }
}

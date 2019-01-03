namespace ExcelWriter
{
    /// <summary>
    /// When associated with a data element, determines the index of the series in the template chart that will be used to represent this data.
    /// </summary>
    public class ChartBasedOnSeriesIndexOption : ChartOptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChartBasedOnSeriesIndexOption" /> class.
        /// </summary>
        public ChartBasedOnSeriesIndexOption()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartBasedOnSeriesIndexOption" /> class.
        /// </summary>
        public ChartBasedOnSeriesIndexOption(int seriesIndex)
        {
            this.SeriesIndex = seriesIndex;
        }

        /// <summary>
        /// Gets or sets the index of the series in the template chart that will be used to represent this data.
        /// </summary>
        public int SeriesIndex { get; set; }
    }
}

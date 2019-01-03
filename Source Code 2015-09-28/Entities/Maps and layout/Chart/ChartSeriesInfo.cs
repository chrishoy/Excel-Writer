namespace ExcelWriter
{
    using System;
    using System.Windows.Media;

    /// <summary>
    /// Represents information relating to a data element when used for creating chart series.<br/>
    /// </summary>
    internal class ChartSeriesInfo
    {
        #region Private Fields

        private bool suppressSeries;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeriesInfo" /> class based on the supplied data.
        /// </summary>
        /// <param name="sourceData">The data on which this information is based</param>
        public ChartSeriesInfo(object sourceData)
        {
            if (sourceData == null)
            {
                throw new ArgumentNullException("sourceData");
            }

            if (sourceData is IChartMetadata)
            {
                // IChartMetadata is used to determine chart related options first
                this.InitChartSeriesInfo(sourceData as IChartMetadata);
            }
        }

        #endregion Construction

        #region Public Properies

        /// <summary>
        /// Gets a value indicating whether the series should be suppressed for the row.
        /// </summary>
        public bool SuppressSeries 
        {
            get
            {
                return this.suppressSeries;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the row should be used as the Category1 Axis in any charts.
        /// </summary>
        public bool IsCategory1Axis { get; set; }

        /// <summary>
        /// Gets or sets the index of the series in the template chart that should be used as the basis of generating series dynamically.
        /// </summary>
        public int BaseOnChartSeriesIndex { get; set; }

        /// <summary>
        /// Gets or sets a colour to be used for the series.
        /// </summary>
        public Color? SeriesColour { get; set; }

        #endregion Public Properties

        #region Private Helpers

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeriesInfo" /> class.
        /// </summary>
        /// <param name="sourceData">The data which is used to create the series</param>
        private void InitChartSeriesInfo(IChartMetadata sourceData)
        {
            if (sourceData.ChartOptions != null)
            {
                foreach (ChartOptionBase option in sourceData.ChartOptions)
                {
                    if (option is ChartSeriesColourOption)
                    {
                        // Option to control the series colour
                        this.SeriesColour = (option as ChartSeriesColourOption).Colour;
                    }
                    else if (option is ChartCategory1AxisOption)
                    {
                        // Option to control whether this data is to be used as the Category 1 Axis
                        this.IsCategory1Axis = (option as ChartCategory1AxisOption).IsCategory1Axis;
                    }
                    else if (option is ChartBasedOnSeriesIndexOption)
                    {
                        // Option to control which tempalte series to used as the basis of representing this data.
                        this.BaseOnChartSeriesIndex = (option as ChartBasedOnSeriesIndexOption).SeriesIndex;
                    }
                }
            }

            // If we base a series on a series which can not exist in the template, then we
            // obviously don't want to see the series
            if (this.BaseOnChartSeriesIndex < 0)
            {
                this.suppressSeries = true;
            }
        }

        #endregion Private Helpers
    }
}

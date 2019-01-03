namespace ExcelWriter
{
    using System.Windows.Media;

    /// <summary>
    /// When associated with a data element, determines the colour that will be used when it is presented within a chart.
    /// </summary>
    public class ChartSeriesColourOption : ChartOptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeriesColourOption" /> class.
        /// </summary>
        public ChartSeriesColourOption()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeriesColourOption" /> class.
        /// </summary>
        public ChartSeriesColourOption(Color colour)
        {
            this.Colour = colour;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSeriesColourOption" /> class.
        /// </summary>
        public ChartSeriesColourOption(Color? colour)
        {
            this.Colour = colour;
        }

        /// <summary>
        /// Gets or sets the colour that will be used to represent this data entity within a chart.<br/>
        /// If null, then colour will not be explicitly set.
        /// </summary>
        public Color? Colour { get; set; }
    }
}

namespace ExcelWriter
{
    /// <summary>
    /// When associated with a data element, determines whether this data will be excluded from charts.
    /// </summary>
    public class ChartExcludeOption : ChartOptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChartExcludeOption" /> class.
        /// </summary>
        public ChartExcludeOption()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartExcludeOption" /> class.
        /// </summary>
        /// <param name="isCategory1Axis">True if this data element is to be excluded</param>
        public ChartExcludeOption(bool exclude)
        {
            this.Exclude = exclude;
        }

        /// <summary>
        /// Gets or sets a value indicating whether this data will be excluded from charts.
        /// </summary>
        public bool Exclude { get; set; }
    }
}

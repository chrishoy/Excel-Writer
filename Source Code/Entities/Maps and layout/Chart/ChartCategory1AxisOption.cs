namespace ExcelWriter
{
    /// <summary>
    /// When associated with a data element, determines whether this data will be used to represent the 'Category 1 Axis'.
    /// </summary>
    public class ChartCategory1AxisOption : ChartOptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChartCategory1AxisOption" /> class.
        /// </summary>
        public ChartCategory1AxisOption()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartCategory1AxisOption" /> class.
        /// </summary>
        /// <param name="isCategory1Axis">True if this data element represents the 'Category 1 Axis'</param>
        public ChartCategory1AxisOption(bool isCategory1Axis)
        {
            this.IsCategory1Axis = isCategory1Axis;
        }

        /// <summary>
        /// Gets or sets a value indicating whether this data will represent the 'Category 1 Axis'
        /// </summary>
        public bool IsCategory1Axis { get; set; }
    }
}

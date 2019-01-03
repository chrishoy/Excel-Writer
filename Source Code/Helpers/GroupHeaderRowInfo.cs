namespace ExcelWriter
{
    /// <summary>
    /// Represents information about the row where a group header is to be written.
    /// </summary>
    internal class GroupHeaderRowInfo
    {
        /// <summary>
        /// Gets or sets a value indicating the level above the columns that the group will reside
        /// </summary>
        /// <value>
        /// The level.
        /// </value>
        public int Level { get; set; }

        /// <summary>
        /// Gets or sets a value which is calculated from the maximum header height
        /// </summary>
        /// <value>
        /// The height.
        /// </value>
        public double? Height { get; set; }

        /// <summary>
        /// Gets or sets a value which represents the hidden state of the group header row.
        /// </summary>
        /// <value>
        ///   <c>true</c> if hidden; otherwise, <c>false</c>.
        /// </value>
        public bool Hidden { get; set; }
    }
}
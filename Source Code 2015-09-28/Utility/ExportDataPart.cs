namespace ExcelWriter
{
    /// <summary>
    /// Represents an identifiable piece of data which is the source for a section of a report.
    /// </summary>
    public class ExportDataPart
    {
        /// <summary>
        /// Gets or sets the source data to be rendered as an export report.
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// Gets or sets an Id which can be used to identify this data and match it to a report rendering element.
        /// </summary>
        public string PartId { get; set; }
    }
}

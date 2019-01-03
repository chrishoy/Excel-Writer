
namespace ExcelWriter
{
    /// <summary>
    /// Used for copying ranges of data from one excel worksheet to another
    /// </summary>
    public class RangeMapping : VisualMapping
    {
        /// <summary>
        /// The number of rows to copy from the source sheet
        /// </summary>
        public int? RowCount { get; set; }

        /// <summary>
        /// The index of the source sheet to start copying from
        /// </summary>
        public int SourceStartRowIndex { get; set; }

        /// <summary>
        /// The index of the target sheet to start copying to
        /// </summary>
        public int TargetStartRowIndex { get; set; }

        /// <summary>
        /// When true copy data from the presentation sheet, when false
        /// copy data from the data sheet.
        /// Default value is false.
        /// </summary>
        public bool UsePresentationAsSource { get; set; }
    }
}

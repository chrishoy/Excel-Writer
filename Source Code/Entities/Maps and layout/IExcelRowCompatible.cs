namespace ExcelWriter
{
    /// <summary>
    /// A map that implements this interface will influence row creation when exporting.<br/>
    /// This includes row height and visibility.
    /// </summary>
    public interface IExcelRowCompatible
    {
        /// <summary>
        ///  Gets or sets a value that determines the height that this cell should attempt to be when exported to Excel. Null is 'go with the flow'.
        /// </summary>
        double? Height { get; set; }

        /// <summary>
        /// Gets or sets the value that is used to hide the row where this cell resides.
        /// </summary>
        object RowIsHidden { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        int RowSpan { get; set; }
    }
}

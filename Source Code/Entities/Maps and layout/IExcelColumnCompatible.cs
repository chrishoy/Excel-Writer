namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// A map that implements this interface will influence column creation when exporting.<br/>
    /// This includes column width and visibility.
    /// </summary>
    public interface IExcelColumnCompatible
    {
        /// <summary>
        ///  Gets or sets a value that determines the width that this cell should attempt to be when exported to Excel. Null is 'go with the flow'.
        /// </summary>
        double? Width { get; set; }

        /// <summary>
        /// Gets or sets the value that is used to hide the column where this cell resides.
        /// </summary>
        object ColumnIsHidden { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        int ColumnSpan { get; set; }
    }
}

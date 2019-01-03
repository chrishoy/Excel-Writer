namespace ExcelWriter
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents information which enables a table column heading to be mapped to Excel.<br/>
    /// Holds the column and any column header which may have been applied to that column
    /// </summary>
    internal class TableColumnInfo
    {
        #region Private Fields

        private Dictionary<int, TableColumnHeader> columnHeaders;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="TableColumnInfo" /> class.
        /// </summary>
        public TableColumnInfo()
        {
            this.columnHeaders = new Dictionary<int, TableColumnHeader>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the <see cref="TableColumn"/> which this information is for.
        /// </summary>
        public TableColumn Column { get; set; }

        /// <summary>
        /// Gets the width of the column (and header)
        /// </summary>
        public double? Width { get; set; }

        /// <summary>
        /// Gets the number of columns this column spans
        /// </summary>
        public int ColumnSpan { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this column should be spanned to the end of the container.
        /// </summary>
        public bool SpanLastColumn { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this is the last column in the table
        /// </summary>
        public bool IsLastColumn { get; set; }

        /// <summary>
        /// Gets a value indicating whether the column is hidden
        /// </summary>
        public bool Hidden { get; set; }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Gets the <see cref="TableColumnHeader"/> for the specified level.<br/>
        /// </summary>
        /// <param name="level">The requested level</param>
        /// <returns>The <see cref="TableColumnHeader"/> or null if no <see cref="TableColumnHeader"/> at that level.</returns>
        public TableColumnHeader GetColumnHeader(int level)
        {
            if (this.columnHeaders.ContainsKey(level))
            {
                return this.columnHeaders[level];
            }

            return null;
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Assigns a column header to this <see cref="TableColumnInfo"/>.
        /// </summary>
        /// <param name="level">The level this column header is at</param>
        /// <param name="columnHeader">The <see cref="TableColumnHeader"/> to be assigned</param>
        internal void AddHeader(int level, TableColumnHeader columnHeader)
        {
            this.columnHeaders.Add(level, columnHeader);
        }

        #endregion Internal Methods
    }
}

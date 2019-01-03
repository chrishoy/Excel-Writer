namespace ExcelWriter
{
    using System.Windows;

    /// <summary>
    /// Represents a column group over a set of <see cref="TableColumn"/>s.
    /// </summary>
    public class TableColumnHeader
    {
        #region Private Fields

        private object dataContext;
        private object header;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="TableColumnHeader" /> class.
        /// </summary>
        public TableColumnHeader()
        {
            this.HeaderStyleKey = SharedMapStyleKeys.ColumnHeaderBorderExport;
            this.Visibility = System.Windows.Visibility.Visible;
        }

        #endregion Construction

        /// <summary>
        /// Gets or sets the key used to look up a header map style (non-dependency property)
        /// </summary>
        public string HeaderStyleKey { get; set; }

        /// <summary>
        /// Gets or sets the data source for instance.
        /// </summary>
        public object DataContext
        {
            get { return BindingContainer.EvaluateIfRequired(this.dataContext, this.ParentDataContext); }
            set { this.dataContext = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the column number within the table where this column header start to span.
        /// </summary>
        public int Finish { get; set; }

        /// <summary>
        /// Gets or sets the column number within the table up to where this column header spans.
        /// </summary>
        public int Start { get; set; }

        /// <summary>
        ///  Gets or sets the height that this cell should attempt to be when exported to Excel.<br />
        ///  If null, then height is determined by surrounding cells.
        /// </summary>
        public double? Height { get; set; }

        /// <summary>
        /// Gets or sets whether the column header for this column is visible (i.e. row hidden when exported to Excel)
        /// </summary>
        public Visibility Visibility
        {
            get; set; 
        }

        /// <summary>
        /// Gets or sets a value that is displayed above the column of data when exported to Excel
        /// </summary>
        public object Header
        {
            get { return BindingContainer.EvaluateIfRequired(this.header, this.DataContext); }
            set { this.header = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the level which this column header should be assigned to.
        /// </summary>
        public int Level { get; set; }

        #region Internal Properties

        /// <summary>
        /// Gets or sets the parent data context. Used for binding the DataContext itself.
        /// </summary>
        internal object ParentDataContext { get; set; }

        #endregion Internal Properties
    }
}

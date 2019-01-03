namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows.Markup;

    /// <summary>
    /// Represents a XAML definition for the data for a table, chart, or both, which can be exported to an Excel worksheet.
    /// </summary>
    [ContentProperty("Columns")]
    public class TableData : BaseMap, IExcelTableDataPreparable
    {
        #region Private Fields

        private object itemsSource;
        private TableColumnCollection columns;
        private List<TableDataRowInfo> rowData;

        // Used internally to track the data region when written to Excel
        private ExcelMapCoOrdinateContainer dataRegion;

        private bool treatRowAsSeries;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Public ctor. Creates instance of <see cref="TableData"/>
        /// </summary>
        public TableData() : base()
        {
            this.columns = new TableColumnCollection();
            this.rowData = new List<TableDataRowInfo>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Source for the repeating rows in the table (Simple Bindable)
        /// </summary>
        public object ItemsSource
        {
            get { return BindingContainer.EvaluateIfRequired(this.itemsSource, this.DataContext); }
            set { this.itemsSource = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Collection of columns for the table.
        /// </summary>
        public TableColumnCollection Columns
        {
            get { return this.columns; }
            set { this.columns = value; }
        }

        /// <summary>
        /// Default is false, each column forms a series
        /// When true each row forms a series
        /// </summary>
        public bool TreatRowAsSeries
        {
            get { return this.treatRowAsSeries; }
            set { this.treatRowAsSeries = value; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets a value which indicates that this <see cref="TableData"/> is not visual, i.e. is never written into Exel
        /// </summary>
        internal override bool IsVisual
        {
            get { return false; }
        }

        /// <summary>
        /// Maintains an internal list of the elements that the rows of this <see cref="TableData"/> represent.
        /// </summary>
        internal List<TableDataRowInfo> RowData
        {
            get { return this.rowData; }
        }

        /// <summary>
        /// Gets or sets a <see cref="ExcelMapCoOrdinateContainer">Table data region map</see> used internally to track
        /// the area in the Excel worksheet where this entity is written.
        /// </summary>
        internal ExcelMapCoOrdinateContainer MapContainer
        {
            get { return this.dataRegion; }
            set { this.dataRegion = value; }
        }

        #endregion Internal Properties

        #region Internal Methods

        /// <summary>
        /// Add a row to an internal list of the elements that this <see cref="TableData"/> represents. 
        /// </summary>
        /// <param name="item"></param>
        internal void AddRow(object item)
        {
            this.rowData.Add(new TableDataRowInfo(item));
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="Map"/> in this <see cref="Map"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="Map"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T) return (T)(BaseMap)this;
            return null;
            //NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// which has a specified key. This includes this instance.
        /// </summary>
        /// <param name="key">The key of the typed item that we require</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>(string key)
        {
            if (this is T && this.Key == key)
            {
                return (T)(BaseMap)this;
            }

            return null;

            // NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="TableData"/><br/>.
        /// This includes this instance.
        /// </summary>
        /// <param name="list">The list to be updated</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find</typeparam>
        internal override void AddDescendentsOfType<T>(ref List<T> list)
        {
            if (this is T)
            {
                list.Add((T)(BaseMap)this);
            }

            // For the moment, don't go any lower
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("TableData");
        }

        #endregion Internal Methods

    }
}

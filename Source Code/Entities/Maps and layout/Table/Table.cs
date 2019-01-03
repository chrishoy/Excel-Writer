namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows.Markup;

    /// <summary>
    /// Represents a XAML definition for a table that can be exported to an Excel worksheet.
    /// </summary>
    [ContentProperty("Columns")]
    public class Table : BaseMap, IExcelTablePreparable
    {
        #region Private Fields

        private PropertyCollection properties;
        private TableColumnHeaderCollection columnHeaders;
        private TableColumnCollection columns;
        private TableData tableData;
        private object tableDataKey;

        private string headerStyleKey;
        private string footerStyleKey;
        private string subHeaderStyleKey;
        private string subFooterStyleKey;
        private string rowStyleSelectorKey;

        // Below are 'Bindable'
        private object dataRegionDefinedName;
        private object header;
        private object footer;
        private object subHeader;
        private object subFooter;
        private bool hideColumnsHeader;
        private object itemsSource;
        private object defaultRowHeight;

        private string cellStyleKey;
        private string cellStyleSelectorKey;

        // Below are not 'Bindable'
        private bool spanLastColumn;
        private bool padLastColumn;
        private List<ExcelChartOptions> excelChartOptionsList;
        private MapCollection items;

        #endregion Private Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Table" /> class.
        /// </summary>
        public Table() : base()
        {
            // Set default values
            this.headerStyleKey = SharedMapStyleKeys.HeaderExport;
            this.subHeaderStyleKey = SharedMapStyleKeys.SubHeaderExport;
            this.footerStyleKey = SharedMapStyleKeys.FooterExport;
            this.subFooterStyleKey = SharedMapStyleKeys.SubFooterExport;

            this.properties = new PropertyCollection();
            this.columnHeaders = new TableColumnHeaderCollection();
            this.columns = new TableColumnCollection();
            this.excelChartOptionsList = new List<ExcelChartOptions>();
            this.items = new MapCollection();
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets a style key.
        /// </summary>
        public string CellStyleKey
        {
            get { return this.cellStyleKey; }
            set { this.cellStyleKey = value; }
        }

        /// <summary>
        /// Gets or sets the key to a CellStyleSelector that can be used to choose CellStyleKeys at runtime
        /// </summary>
        public string CellStyleSelectorKey
        {
            get { return this.cellStyleSelectorKey; }
            set { this.cellStyleSelectorKey = value; }
        }

        /// <summary>
        /// Gets or sets a value which will be used to set a 'Defined Name' when exported to Excel. (This is a property that can be used with binding)<br/>
        /// This region includes the main column headers, but excludes any group headers.<br/>
        /// Each column will also be assigned a 'Defined Name' which is the DataRegionDefinedName + '_' + ColumnName.
        /// </summary>
        public object DataRegionDefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.dataRegionDefinedName, this.DataContext); }
            set { this.dataRegionDefinedName = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a style. Replacement for FooterStyle (non-dependency property)
        /// </summary>
        public string FooterStyleKey
        {
            get { return this.footerStyleKey; }
            set { this.footerStyleKey = value; }
        }

        /// <summary>
        /// Gets or sets a style. Replacement for HeaderStyle (non-dependency property)
        /// </summary>
        public string HeaderStyleKey
        {
            get { return this.headerStyleKey; }
            set { this.headerStyleKey = value; }
        }

        /// <summary>
        /// Gets or sets a style. Replacement for SubFooterStyle (non-dependency property)
        /// </summary>
        public string SubFooterStyleKey
        {
            get { return this.subFooterStyleKey; }
            set { this.subFooterStyleKey = value; }
        }

        /// <summary>
        /// Gets or sets a style. Replacement for SubHeaderStyle (non-dependency property)
        /// </summary>
        public string SubHeaderStyleKey
        {
            get { return this.subHeaderStyleKey; }
            set { this.subHeaderStyleKey = value; }
        }

        /// <summary>
        /// Gets or sets the key to a CellStyleSelector that will be used to choose CellStyleKeys at runtime for all cells in a table row.
        /// </summary>
        public string RowStyleSelectorKey
        {
            get { return this.rowStyleSelectorKey; }
            set { this.rowStyleSelectorKey = value; }
        }

        /// <summary>
        /// Gets or sets a value which determines the header for the table (This is a property that can be used with binding)
        /// </summary>
        public object Header
        {
            get { return BindingContainer.EvaluateIfRequired(this.header, this.DataContext); }
            set { this.header = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the footer for the table (This is a property that can be used with binding)
        /// </summary>
        public object Footer
        {
            get { return BindingContainer.EvaluateIfRequired(this.footer, this.DataContext); }
            set { this.footer = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the sub-header for the table (This is a property that can be used with binding)
        /// </summary>
        public object SubHeader
        {
            get { return BindingContainer.EvaluateIfRequired(this.subHeader, this.DataContext); }
            set { this.subHeader = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the sub-footer for the table (This is a property that can be used with binding)
        /// </summary>
        public object SubFooter
        {
            get { return BindingContainer.EvaluateIfRequired(this.subFooter, this.DataContext); }
            set { this.subFooter = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a default row height for table rows within the table (This is a property that can be used with binding)
        /// </summary>
        public object DefaultRowHeight
        {
            get { return BindingContainer.EvaluateIfRequired(this.defaultRowHeight, this.DataContext); }
            set { this.defaultRowHeight = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets a collection of <see cref="Property"/>s which will be written above the <see cref="Table"/>
        /// </summary>
        public PropertyCollection Properties
        {
            get { return this.properties; }
        }

        /// <summary>
        /// Gets a collection of column headers which may span across columns.
        /// </summary>
        public TableColumnHeaderCollection ColumnHeaders
        {
            get { return this.columnHeaders; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether columns header directly above the column data (as opposed to the ColumnHeaders which are above) are to be suppressed.<br/>
        /// When exporting to Excel, this sets the Hidden property on the single row containing the columns header to True.<br/>
        /// This is a dependency property.
        /// </summary>
        public bool HideColumnsHeader
        {
            get { return this.hideColumnsHeader; }
            set { this.hideColumnsHeader = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the last column cells are spanned (merged) to the end of the containing element.<br/>
        /// This takes precedence over PadLastColumn.
        /// </summary>
        public bool SpanLastColumn
        {
            get { return this.spanLastColumn; }
            set { this.spanLastColumn = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a padding column should be added to
        /// the table if there are fewer columns than would otherwise fill the containg element.<br/>
        /// The padding column will merge up the the end of the containg element.<br/>
        /// NB! SpanLastColumn takes precedence over PadLastColumn.
        /// </summary>
        public bool PadLastColumn
        {
            get { return this.padLastColumn; }
            set { this.padLastColumn = value; }
        }

        /// <summary>
        /// Gets or sets a <see cref="TableColumnCollection">collection</see> of columns for the table.
        /// </summary>
        public TableColumnCollection Columns
        {
            get { return this.columns; }
            set { this.columns = value; }
        }

        /// <summary>
        /// Gets or sets a source for the repeating rows in the table (This is a property that can be used with binding)
        /// </summary>
        public object ItemsSource
        {
            get { return BindingContainer.EvaluateIfRequired(this.itemsSource, this.DataContext); }
            set { this.itemsSource = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a <see cref="TableData"/> which represents the data behind the table.
        /// </summary>
        public TableData TableData
        {
            get { return this.tableData; }
            set { this.tableData = value; }
        }

        /// <summary>
        /// Key to the table data behind the table.</br>
        /// NB! This can not be used in conjunction with explicit or implicit TableData defined using Columns.
        /// </summary>
        public object TableDataKey
        {
            get { return BindingContainer.EvaluateIfRequired(this.tableDataKey, this.DataContext); }
            set { this.tableDataKey = BindingContainer.CreateIfRequired(value); }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// A collection of Maps to process (may be added when CellTemplates encountered...
        /// If a CellTemplateMapKey is specified then a <see cref="ContentControl"/> will be added to host the cell template.<br/>
        /// This is later processed by the Export Processing routines.
        /// </summary>
        internal MapCollection Items
        {
            get { return this.items; }
            set { this.items = value; }
        }

        /// <summary>
        /// Creates and populates the default TableData property from the ItemsSource and DataContext
        /// </summary>
        /// <param name="hostTableAlreadyPrepared">True if the host <see cref="Table"/> has already been prepared.</param>
        internal void CreateDefaultTableData(bool hostTableAlreadyPrepared)
        {
            this.tableData = new TableData
            {
                DataContext = BindingContainer.GetSourceBindingOrValue(this.DataContext),
                ItemsSource = BindingContainer.GetSourceBindingOrValue(this.ItemsSource),
                Columns = this.Columns,
                Prepared = hostTableAlreadyPrepared,
            };
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="Map"/> in this <see cref="Map"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="Map"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T)
            {
                return (T)(BaseMap)this;
            }

            // TableData is derived from BaseMap to try this...
            if (this.TableData != null)
            {
                BaseMap item = this.TableData.FirstDescendentOfType<T>();
                if (item != null)
                {
                    return (T)(BaseMap)item;
                }
            }

            // The Properties of a table are derived from BaseMap, so I suppose we should test these...
            if (this.properties == null || this.properties.Count == 0)
            {
                return null;
            }

            for (int i = 0; i < this.properties.Count; i++)
            {
                BaseMap item = this.properties[i].FirstDescendentOfType<T>();
                if (item != null)
                {
                    return (T)(BaseMap)item;
                }
            }

            // NB! For the moment, we will ignore anything lower.
            return null;
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

            // TableData is derived from BaseMap so try this...
            if (this.TableData != null)
            {
                BaseMap item = this.TableData.FirstDescendentOfType<T>(key);
                if (item != null)
                {
                    return (T)(BaseMap)item;
                }
            }

            // The Properties of a table are derived from BaseMap, so I suppose we should test these...
            if (this.properties != null && this.properties.Count > 0)
            {
                for (int i = 0; i < this.properties.Count; i++)
                {
                    BaseMap item = this.properties[i].FirstDescendentOfType<T>(key);
                    if (item != null)
                    {
                        return (T)(BaseMap)item;
                    }
                }
            }

            // Process any Item which may have been dynamically added.
            foreach (BaseMap i in this.Items)
            {
                BaseMap item = i.FirstDescendentOfType<T>(key);
                if (item != null)
                {
                    return (T)(BaseMap)item;
                }
            }

            // NB! For the moment, we will ignore anything lower.
            return null;
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="Table"/><br/>.
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

            // TableData is derived from BaseMap to try this...
            if (this.TableData != null)
            {
                this.TableData.AddDescendentsOfType<T>(ref list);
            }

            // The Properties of a table are derived from BaseMap, so I suppose we should test these...
            if (this.properties != null || this.properties.Count == 0)
            {
                for (int i = 0; i < this.properties.Count; i++)
                {
                    this.properties[i].AddDescendentsOfType<T>(ref list);
                }
            }

            // Process any Item which may have been dynamically added.
            if (this.items != null && this.items.Count > 0)
            {
                for (int i = 0; i < this.items.Count; i++)
                {
                    this.items[i].AddDescendentsOfType<T>(ref list);
                }
            }

            // For the moment, don't go any lower
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Table");
        }

        #endregion Internal Methods
    }
}

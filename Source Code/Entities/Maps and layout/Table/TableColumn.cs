namespace ExcelWriter
{
    /// <summary>
    /// Represents a column of data when exported to Excel.
    /// </summary>
    public class TableColumn : IChartMetadata, IExcelColumnCompatible
    {
        #region Private Fields

        private readonly ChartOptions chartOptions;

        private int columnSpan;
        private int rowSpan;
        private string cellTemplateMapKey;

        // Bindable
        private object header;
        private object dataContext;
        private object columnIsHidden;

        // Used internally to track the data region when written to Excel
        private ExcelMapCoOrdinateContainer dataRegion;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="TableColumn" /> class.
        /// </summary>
        public TableColumn()
        {
            this.CellStyleKey = SharedMapStyleKeys.CellExport;
            this.HeaderStyleKey = SharedMapStyleKeys.ColumnHeaderExport;
            this.chartOptions = new ChartOptions();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the data source for instance.
        /// </summary>
        public object DataContext
        {
            get { return BindingContainer.EvaluateIfRequired(this.dataContext, this.ParentDataContext); }
            set { this.dataContext = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the value that is used to hide the excel worksheet column where this column resides.
        /// </summary>
        public object ColumnIsHidden
        {
            get { return BindingContainer.EvaluateIfRequired(this.columnIsHidden, this.DataContext); }
            set { this.columnIsHidden = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value which determines the width that this cell should attempt to be when exported to Excel.
        /// </summary>
        public double? Width { get; set; }

        /// <summary>
        ///  Gets or sets a value which determines the height that this cell should attempt to be when exported to Excel.
        /// </summary>
        public double? Height { get; set; }

        /// <summary>
        /// Gets or sets a value that is displayed above the column of data when exported to Excel
        /// </summary>
        public object Header 
        {
            get { return BindingContainer.EvaluateIfRequired(this.header, this.DataContext); }
            set { this.header = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        public int ColumnSpan
        {
            get { return this.columnSpan; }
            set { this.columnSpan = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating the number of Excel rows that the header of this column should span.
        /// </summary>
        public int RowSpan
        {
            get { return this.rowSpan; }
            set { this.rowSpan = value; }
        }

        /// <summary>
        /// Gets a list of options that will be used when representing this data within a chart.
        /// </summary>
        public ChartOptions ChartOptions
        {
            get { return this.chartOptions; }
        }

        /// <summary>
        /// Gets or sets a style key.
        /// </summary>
        public string CellStyleKey { get; set; }

        /// <summary>
        /// Gets or sets the key to a CellStyleSelector that can be used to choose CellStyleKeys at runtime
        /// </summary>
        public string CellStyleSelectorKey { get; set; }
        
        /// <summary>
        /// Gets or sets a HeaderStyle
        /// </summary>
        public string HeaderStyleKey { get; set; }

        /// <summary>
        /// Gets or sets the key to a CellStyleSelector that can be used to choose HeaderStyleKey at runtime
        /// </summary>
        public string HeaderStyleSelectorKey { get; set; }

        /// <summary>
        /// Gets or sets the member of ItemsSource that will be displayed when exported to Excel
        /// </summary>
        public string DisplayMember { get; set; }

        ///// <summary>
        ///// Used in conjunction with the ItemsSource
        ///// Foreach IDataPart in the ItemsSource a new map will be supplied based on this ItemTemplateMap
        ///// </summary>
        //public ContentControl CellTemplateMap
        //{
        //    get { return this.cellTemplateMap; }
        //    set { this.cellTemplateMap = value; }
        //}

        /// <summary>
        /// Used in conjunction with ItemsSource and .<br/>
        /// For Each item in ItemsSource, a new map will be created and supplied from the re-usable Maps collection of the package.
        /// </summary>
        public string CellTemplateMapKey
        {
            get { return this.cellTemplateMapKey; }
            set { this.cellTemplateMapKey = value; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets or sets a <see cref="ExcelMapCoOrdinateContainer">Column map</see> used internally to track the area in the Excel worksheet where data is written.
        /// </summary>
        internal ExcelMapCoOrdinateContainer DataRegion
        {
            get { return this.dataRegion; }
            set { this.dataRegion = value; }
        }

        /// <summary>
        /// Gets or sets the parent data context. Used for binding the DataContext itself.
        /// </summary>
        internal object ParentDataContext { get; set; }

        #endregion Internal Properties
    }
}

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Windows;

    using OpenXml.Excel.Model;
    
    /// <summary>
    /// Represents a XAML definition for a chart that can be exported to an Excel worksheet.
    /// A <see cref="Chart"/> must be based on a <see cref="ChartTemplate"/>.
    /// </summary>
    public class Chart : BaseMap, IExcelColumnCompatible, IExcelRowCompatible
    {
        #region Private Fields

        private TableData tableData;

        private double? width;          // Defaulted to null
        private double? height;         // Defaulted to null
        private object definedName;
        private object title;
        private bool spanLastColumn;
        private int columnSpan;
        private bool spanLastRow;
        private int rowSpan;

        private string cellStyleKey;
        private string cellStyleSelectorKey;

        private Visibility visibility;

        private ExcelMapCoOrdinatePlaceholder mapPlaceholder;
        private List<ExcelChartOptions> optionsList;

        // Bindable
        private object chartTemplateKey;
        private object tableDataKey;
        private object rowIsHidden;
        private object columnIsHidden;

        #endregion Private Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Chart" /> class.
        /// </summary>
        public Chart() : base()
        {
            // Set default values
            this.optionsList = new List<ExcelChartOptions>();
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
        /// Gets or sets the name of the <see cref="ChartTemplate"/> resource which will be used to derive this chart.<br/>
        /// This a Bindable property.
        /// </summary>
        public object ChartTemplateKey
        {
            get { return BindingContainer.EvaluateIfRequired(this.chartTemplateKey, this.DataContext); }
            set { this.chartTemplateKey = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a list of options that can be applied to this chart.
        /// </summary>
        [Obsolete("Not yet implemented")]
        public List<ExcelChartOptions> OptionsList
        {
            get { return this.optionsList; }
            set { this.optionsList = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the last cell (i.e. this) is to be spanned (merged) to the end of the containing element.
        /// </summary>
        public bool SpanLastColumn
        {
            get { return this.spanLastColumn; }
            set { this.spanLastColumn = value; }
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
        /// Gets or sets a value indicating whether the last cell (i.e. this) is to be spanned (merged) to the end of the containing element.
        /// </summary>
        public bool SpanLastRow
        {
            get { return this.spanLastRow; }
            set { this.spanLastRow = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        public int RowSpan
        {
            get { return this.rowSpan; }
            set { this.rowSpan = value; }
        }

        /// <summary>
        /// Gets or sets whether the row where this cell resides is visible (i.e. Hidden in Excel).<br/>
        /// It was a lame attempt at hiding rows. Use RowIsHidden and ColumnIsHidden instead.
        /// </summary>
        [Obsolete("This property has been superseded by RowIsHidden and ColumnIsHidden")]
        public Visibility Visibility
        {
            get
            {
                return this.visibility;
            }

            set
            {
                this.visibility = value;

                // If you set Visibility to Hidden or Collapsed, and RowIsHidden is not being used, then setRowIsHidden to True
                if (value != System.Windows.Visibility.Visible && this.rowIsHidden == null)
                {
                    this.rowIsHidden = true;
                }
            }
        }

        /// <summary>
        /// Gets or sets the value that is used to hide the row where this cell resides.
        /// </summary>
        public object RowIsHidden
        {
            get { return BindingContainer.EvaluateIfRequired(this.rowIsHidden, this.DataContext); }
            set { this.rowIsHidden = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets the value that is used to hide the column where this cell resides.
        /// </summary>
        public object ColumnIsHidden
        {
            get { return BindingContainer.EvaluateIfRequired(this.columnIsHidden, this.DataContext); }
            set { this.columnIsHidden = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        ///  Gets or sets a value that determines the width that this cell should attempt to be when exported to Excel. Null is 'go with the flow'.
        /// </summary>
        public double? Width
        {
            get { return this.width; }
            set { this.width = value; }
        }

        /// <summary>
        ///  Gets or sets a value that determines the height that this cell should attempt to be when exported to Excel. Null is 'go with the flow'.
        /// </summary>
        public double? Height
        {
            get { return this.height; }
            set { this.height = value; }
        }

        /// <summary>
        /// Gets or sets a value that determines a 'Defined Name' that the cell/cells where this chart resides will be given when exported to Excel (This is a property that can be used with binding)
        /// </summary>
        public object DefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.definedName, this.DataContext); }
            set { this.definedName = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value which will be set as the title for the chart (if defined in the chart template) when exported Excel (This is a property that can be used with binding)
        /// </summary>
        public object Title
        {
            get { return BindingContainer.EvaluateIfRequired(this.title, this.DataContext); }
            set { this.title = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a reference to the TableData that will be supplying data to this chart. <br/>
        /// This will be internally set when the chart is being processed.
        /// </summary>
        public TableData TableData
        {
            get { return this.tableData; }
            set { this.tableData = value; }
        }

        /// <summary>
        /// Gets or sets a key which is used to look up the table data behind the chart.
        /// </summary>
        public object TableDataKey
        {
            get { return BindingContainer.EvaluateIfRequired(this.tableDataKey, this.DataContext); }
            set { this.tableDataKey = BindingContainer.CreateIfRequired(value); }
        }

        #endregion

        #region Internal Properties

        /// <summary>
        /// Gets or sets a <see cref="ExcelMapCoOrdinateContainer">Cell map</see> used internally to track
        /// the area in the Excel worksheet where this chart is written.
        /// </summary>
        internal ExcelMapCoOrdinatePlaceholder MapPlaceholder
        {
            get { return this.mapPlaceholder; }
            set { this.mapPlaceholder = value; }
        }

        private object templateWorksheet;

        /// <summary>
        /// Gets or sets a reference to the OpenXml <see cref="DocumentFormat.OpenXml.Packaging.ChartPart"/> in the
        /// template worksheet, from which charts will be cloned.
        /// </summary>
        internal object TemplateWorksheet
        {
            get { return this.templateWorksheet; }
            set { this.templateWorksheet = value; }
        }

        #endregion Internal Properties

        #region Internal Methods

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
            // if (this.TableData != null) return this.TableData.FirstDescendentOfType<T>();

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

            // TableData is derived from BaseMap to try this...
            // if (this.TableData != null) return this.TableData.FirstDescendentOfType<T>(key);

            // NB! For the moment, we will ignore anything lower.
            return null;
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="Cell"/><br/>.
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
            // if (this.TableData != null) return this.TableData.AddDescendentsOfType<T>(ref list);

            // NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Chart");
        }

        #endregion Internal Methods
    }
}

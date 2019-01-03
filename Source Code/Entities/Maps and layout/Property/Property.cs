namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows;
    using System;

    /// <summary>
    /// Represents a XAML definition for a horizontal pair or cells, one being the Header,
    /// the other the Value, that can be exported into an Excel worksheet.
    /// </summary>
    public class Property : BaseMap, IExcelRowCompatible
    {
        #region Private Fields

        private string headerStyleKey;
        private string headerStyleSelectorKey;
        private bool spanLastColumn;
        private double? height;         // Defaulted to null
        private int rowSpan;

        private Visibility visibility;
        private double? cellWidth;

        private string cellStyleKey;
        private string cellStyleSelectorKey;

        // Bindable
        private object header;
        private object value;
        private object rowIsHidden;
        private object definedName;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Default ctor.
        /// </summary>
        public Property()
            : base()
        {
            this.CellStyleKey = SharedMapStyleKeys.CellExport;

            this.headerStyleKey = SharedMapStyleKeys.PropertyHeaderExport;
            this.visibility = Visibility.Visible;
        }

        #endregion Construction

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
        /// Get/set whether the cell for this column is visible (i.e. row hidden when exported to Excel)
        /// </summary>
        [Obsolete("This property has been superceded by RowIsHidden")]
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
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        public int RowSpan
        {
            get { return this.rowSpan; }
            set { this.rowSpan = value; }
        }

        /// <summary>
        ///  Gets or sets a value that determines the height that the cells should attempt to be when exported to Excel. Null is 'go with the flow'.
        /// </summary>
        public double? Height
        {
            get { return this.height; }
            set { this.height = value; }
        }

        /// <summary>
        /// If specified, determines the width that the header and value cells should attempt to be when exported to Excel.
        /// </summary>
        public double? CellWidth
        {
            get { return this.cellWidth; }
            set { this.cellWidth = value; }
        }

        /// <summary>
        /// Replacement for CellStyle (non-dependency property)
        /// </summary>
        public string HeaderStyleKey
        {
            get { return this.headerStyleKey; }
            set { this.headerStyleKey = value; }
        }

        /// <summary>
        /// The key to a CellStyleSelector that can be used to choose CellStyleKeys at runtime
        /// </summary>
        public string HeaderStyleSelectorKey 
        {
            get { return this.headerStyleSelectorKey; }
            set { this.headerStyleSelectorKey = value; }
        }

        /// <summary>
        /// Get the Header for this instance, either derived from the HeaderPath, which uses the DataContext, or explicitly.
        /// NB! Setting this property explicitly takes precendece!
        /// </summary>
        public object Header
        {
            get { return BindingContainer.EvaluateIfRequired(this.header, this.DataContext); }
            set { this.header = BindingContainer.CreateIfRequired(value); }            
        }

        /// <summary>
        /// Gets or sets the value for this instance (This is a property that can be used with binding)
        /// </summary>
        public object Value
        {
            get { return BindingContainer.EvaluateIfRequired(this.value, this.DataContext); }
            set { this.value = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value that determins a 'Defined Name' that this cell will be given when exported to Excel (This is a property that can be used with binding)
        /// </summary>
        public object DefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.definedName, this.DataContext); }
            set { this.definedName = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the last cell (i.e. this) is to be spanned (merged) to the end of the containing element.
        /// </summary>
        public bool SpanLastColumn
        {
            get { return this.spanLastColumn; }
            set { this.spanLastColumn = value; }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="Property"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T) return (T)(BaseMap)this;
            return null;
            //NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="Property"/><br/>
        /// which has a specified key. This includes this instance.
        /// </summary>
        /// <param name="key">The key of the typed item that we require</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>(string key)
        {
            if (this is T && this.Key == key) return (T)(BaseMap)this;
            return null;
            //NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="Property"/><br/>.
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

            // NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Property");
        }

        #endregion Internal Methods
    }
}

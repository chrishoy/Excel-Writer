namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Windows;

    /// <summary>
    /// Represents the base class for a XAML definition for an picture, shape or chart<br/>
    /// that can be exported to an Excel worksheet and subsequently positioned.
    /// </summary>
    public abstract class PositionableMap : BaseMap, IExcelColumnCompatible, IExcelRowCompatible
    {
        #region Private Fields

        private double? width;          // Defaulted to null
        private double? height;         // Defaulted to null
        private bool spanLastColumn;
        private int columnSpan;
        private bool spanLastRow;
        private int rowSpan;
        private string cellStyleKey;
        private string cellStyleSelectorKey;

        private Visibility visibility;

        private ExcelMapCoOrdinatePlaceholder mapPlaceholder;
        private Placement placement;

        // Bindable
        private object definedName;
        private object rowIsHidden;
        private object columnIsHidden;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Constructor
        /// </summary>
        public PositionableMap()
        {
            this.placement = new Placement();
        }

        #endregion Constuction

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
        /// Gets or sets a value indicating the number of Excel cells that this cell should span.
        /// </summary>
        public int RowSpan
        {
            get { return this.rowSpan; }
            set { this.rowSpan = value; }
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
        /// Gets or sets a value that determines a 'Defined Name' that the cell/cells where this picture resides will be given when exported to Excel (This is a property that can be used with binding)
        /// </summary>
        public object DefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.definedName, this.DataContext); }
            set { this.definedName = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets size and position information on the element relative to the cell to which this element is anchored.
        /// </summary>
        public Placement Placement
        {
            get { return this.placement; }
            set { this.placement = value; }
        }

        #endregion

        #region Internal Properties

        /// <summary>
        /// Gets or sets a <see cref="ExcelMapCoOrdinateContainer">Cell map</see> used internally to track
        /// the area in the Excel worksheet where this picture is written.
        /// </summary>
        internal ExcelMapCoOrdinatePlaceholder MapPlaceholder
        {
            get { return this.mapPlaceholder; }
            set { this.mapPlaceholder = value; }
        }

        #endregion Internal Properties

        #region Internal Methods

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T)
            {
                return (T)(BaseMap)this;
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

            // NB! For the moment, we will ignore anything lower.
            return null;
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="Picture"/><br/>.
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
            return base.GetContainerTypeWithKey("Picture");
        }

        #endregion Internal Methods
    }
}

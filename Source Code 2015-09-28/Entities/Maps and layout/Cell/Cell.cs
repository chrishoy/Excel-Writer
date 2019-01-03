namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows.Markup;

    /// <summary>
    /// Represents a XAML definition for a single cell that can be exported to an Excel Worksheet.
    /// </summary>
    [ContentProperty("Value")]
    public class Cell : BaseMap, IExcelColumnCompatible, IExcelRowCompatible
    {
        #region Private Fields

        private double? width;          // Defaulted to null
        private double? height;         // Defaulted to null
        private object definedName;
        private bool spanLastColumn;
        private int columnSpan;
        private bool spanLastRow;
        private int rowSpan;

        private string cellStyleKey;
        private string cellStyleSelectorKey;

        private object rowIsHidden;
        private object columnIsHidden;

        // Bindable
        private object value;

        #endregion Private Fields

        #region Construction

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
        /// Gets or sets the value for this instance (This is a property that can be used with binding)
        /// </summary>
        public object Value
        {
            get { return BindingContainer.EvaluateIfRequired(this.value, this.DataContext); }
            set { this.value = BindingContainer.CreateIfRequired(value); }
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
        /// Gets or sets a value that determins a 'Defined Name' that this cell will be given when exported to Excel (This is a property that can be used with binding)
        /// </summary>
        public object DefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.definedName, this.DataContext); }
            set { this.definedName = BindingContainer.CreateIfRequired(value); }
        }

        #endregion Public Properties

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

            // This is the lowest level element...
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

            // This is the lowest level element...
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

            // This is the lowest level element...
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Cell");
        }

        #endregion Internal Methods
    }
}

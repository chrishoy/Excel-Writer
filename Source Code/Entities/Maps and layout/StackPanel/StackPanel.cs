namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows.Controls;
    using System.Windows.Markup;

    /// <summary>
    /// Provides a mechanism for stacking cells, table, properties, further <see cref="BaseMap"/>s for exporting to Excel.
    /// </summary>
    [ContentProperty("Items")]
    public class StackPanel : BaseMap
    {
        #region Private Fields

        private bool spanLastRow;
        private bool spanLastColumn;
        private object definedName;
        private MapCollection items;
        private string itemTemplateMapKey;
        //private Map itemTemplateMap;
        private Orientation orientation;
        private object itemsSource;

        private string cellStyleKey;
        private string cellStyleSelectorKey;

        #endregion Private Fields

        #region Construction

        public StackPanel()
        {
            this.items = new MapCollection();
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
        /// Gets or sets a value indicating whether the last cell (i.e. this) is to be spanned (merged) to the end of the containing element.
        /// </summary>
        public bool SpanLastColumn
        {
            get { return this.spanLastColumn; }
            set { this.spanLastColumn = value; }
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
        /// Gets or sets a value that determins a 'Defined Name' that this cell will be given when exported to Excel (This is a property that can be used with binding)
        /// </summary>
        public object DefinedName
        {
            get { return BindingContainer.EvaluateIfRequired(this.definedName, this.DataContext); }
            set { this.definedName = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// A collection of Maps to process
        /// If an ItemsSource and ItemTemplateMapKey and  are specified then they will be used instead of ExcelMappings
        /// </summary>
        public MapCollection Items
        {
            get { return this.items; }
            set { this.items = value; }
        }

        ///// <summary>
        ///// Used in conjunction with the ItemsSource
        ///// Foreach IDataPart in the ItemsSource a new map will be supplied based on this ItemTemplateMap
        ///// </summary>
        //public ContentControl ItemTemplateMap
        //{
        //    get { return this.itemTemplateMap; }
        //    set { this.itemTemplateMap = value; }
        //}

        /// <summary>
        /// Used in conjunction with ItemsSource.<br/>
        /// For Each item in ItemsSource, a new map will be created and supplied from the re-usable Maps collection of the package.
        /// </summary>
        public string ItemTemplateMapKey
        {
            get { return this.itemTemplateMapKey; }
            set { this.itemTemplateMapKey = value; }
        }

        /// <summary>
        /// Gets or sets the orientation.
        /// </summary>
        /// <value>
        /// The orientation.
        /// </value>
        public Orientation Orientation
        {
            get { return this.orientation; }
            set { this.orientation = value; }
        }

        /// <summary>
        /// Gets or sets the source of the items in this <see cref="StackPanel"/>.</br>
        /// Used in conjunction with the ItemTemplateMapKey, for each IDataPart in this collection a new <see cref="BaseMap"/> derived entity
        /// will be processed based on the ItemTemplateMap
        /// </summary>
        public object ItemsSource
        {
            get { return BindingContainer.EvaluateIfRequired(this.itemsSource, this.DataContext); }
            set { this.itemsSource = BindingContainer.CreateIfRequired(value); }
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
            if (this is T) return (T)(BaseMap)this;

            // The Properties of a table are derived from Map, so I suppose we should test these...
            if (this.items == null || this.items.Count == 0) return null;
            for (int i = 0; i < this.items.Count; i++)
            {
                BaseMap item = this.items[i].FirstDescendentOfType<T>();
                if (item != null) return (T)(BaseMap)item;
            }
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
            if (this is T && this.Key == key) return (T)(BaseMap)this;

            // The items in this stack panel derive from map so test these...
            if (this.items == null || this.items.Count == 0) return null;
            for (int i = 0; i < this.items.Count; i++)
            {
                BaseMap item = this.items[i].FirstDescendentOfType<T>(key);
                if (item != null) return (T)(BaseMap)item;
            }
            return null;
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="StackPanel"/><br/>.
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

            // The items in this stack panel derive from map so test/add these...
            if (this.items != null)
            {
                for (int i = 0; i < this.items.Count; i++)
                {
                    this.items[i].AddDescendentsOfType<T>(ref list);
                }
            }
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("StackPanel");
        }

        #endregion Internal Methods
    }
}

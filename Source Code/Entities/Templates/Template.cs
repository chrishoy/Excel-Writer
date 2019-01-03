namespace ExcelWriter
{
    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.Windows.Markup;

    /// <summary>
    /// The ExcelTemplate provides a set of instructions for writing data to excel.
    /// Either as a dump of raw data or in a predefined format, ie. chart or table.
    /// </summary>
    [ContentProperty("Content")]
    public sealed class Template : BaseMap
    {
        #region Private Fields

        /// <summary>
        /// The package
        /// </summary>
        private ExcelTemplatePackage package;
        /// <summary>
        /// The map styles
        /// </summary>
        private StylesCollection mapStyles;
        /// <summary>
        /// The maps
        /// </summary>
        private MapCollection maps;
        /// <summary>
        /// The cell style selectors
        /// </summary>
        private CellStyleSelectorCollection cellStyleSelectors;

        /// <summary>
        /// The title
        /// </summary>
        private object title;

        /// <summary>
        /// The data template sheet
        /// </summary>
        private string dataTemplateSheet;
        /// <summary>
        /// The content
        /// </summary>
        private BaseMap content;
        /// <summary>
        /// The presentation template sheet
        /// </summary>
        private string presentationTemplateSheet;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Default ctor.
        /// </summary>
        public Template()
        { }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// The name of the worksheet in the template that stores the raw unformatted data
        /// </summary>
        /// <value>
        /// The data template sheet.
        /// </value>
        public string DataTemplateSheet
        {
            get { return this.dataTemplateSheet; }
            set { this.dataTemplateSheet = value; }
        }

        /// <summary>
        /// When supplied used to write the data to excel
        /// The ExcelMap is an abstract concept that can be one of the following
        /// ExcelCanvas - (to be implemented) but will allow output a collection of ExcelMaps positionally
        /// ExcelStackPanel - outputs a collection of ExcelMaps in either Horizontal or Vertical Orientation
        /// ExcelTableMap - for outputing table of data with 1 or more columns
        /// ExcelPropertyMap - (to be implemented) but will output single cells to excel
        /// </summary>
        /// <value>
        /// The content.
        /// </value>
        public BaseMap Content
        {
            get { return this.content; }
            set { this.content = value; }
        }

        /// <summary>
        /// The name of the worksheet in the template that stores the data in output format
        /// This may be a chart or a formatted table.
        /// This could be the same worksheet at the data.
        /// </summary>
        /// <value>
        /// The presentation template sheet.
        /// </value>
        public string PresentationTemplateSheet
        {
            get { return this.presentationTemplateSheet; }
            set { this.presentationTemplateSheet = value; }
        }

        /// <summary>
        /// Get/Set the title of the document section that this template will export to - This will become (e.g.) the Worksheet Name in Excel.<br />
        /// This is a property that is bindable via a <see cref="BindingExtension" />.
        /// </summary>
        /// <value>
        /// The title.
        /// </value>
        public object Title
        {
            get { return BindingContainer.EvaluateIfRequired(this.title, this.DataContext); }
            set { this.title = BindingContainer.CreateIfRequired(value); }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// PartId and DataTemplateSheet mandatory
        /// More to come?
        /// </summary>
        /// <param name="error">The error.</param>
        /// <returns></returns>
        public bool Valid(out string error)
        {
            error = null;

            StringBuilder s = new StringBuilder();

            // PartId mandatory
            if (string.IsNullOrEmpty(this.TemplateId))
            {
                s.Append("No PartId specified");
                s.Append(Environment.NewLine);
            }

            // DataTemplateSheet mandatory
            if (string.IsNullOrEmpty(this.DataTemplateSheet))
            {
                s.Append("No DataTemplateSheet specified");
                s.Append(Environment.NewLine);
            }

            if (s.Length > 0)
            {
                error = s.ToString();
                return false;
            }
            return true;
        }

        #endregion Public Methods

        #region Internal Fields

        /// <summary>
        /// Gets or sets the cell style selectors.
        /// </summary>
        /// <value>
        /// The cell style selectors.
        /// </value>
        internal CellStyleSelectorCollection CellStyleSelectors
        {
            get { return this.cellStyleSelectors; }
            set { this.cellStyleSelectors = value; }
        }

        /// <summary>
        /// Get/set the <see cref="BaseMap" />s that will be used when looking up re-usable maps. (re-used by Key reference)
        /// </summary>
        /// <value>
        /// The maps.
        /// </value>
        internal MapCollection Maps
        {
            get { return this.maps; }
            set { this.maps = value; }
        }

        /// <summary>
        /// Get/set the <see cref="StylesCollection" /> that will be used when applying non-WPF styles.
        /// </summary>
        /// <value>
        /// The map styles.
        /// </value>
        internal StylesCollection MapStyles
        {
            get { return this.mapStyles; }
            set { this.mapStyles = value; }
        }

        /// <summary>
        /// Get/Set the <see cref="ExcelTemplatePackage" /> that created this template.
        /// </summary>
        /// <value>
        /// The package.
        /// </value>
        internal ExcelTemplatePackage Package
        {
            get { return this.package; }
            set { this.package = value; }
        }

        /// <summary>
        /// The Uri of Xaml in the package
        /// </summary>
        /// <value>
        /// The xaml URI.
        /// </value>
        internal string XamlUri { get; set; }

        #endregion Internal Fields

        #region Internal Methods 

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap" /> in this <see cref="BaseMap" /><br />
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap" /> that we wish to find the first instance of</typeparam>
        /// <returns>
        /// The first instance of type <typeparamref name="T" /> found in the hierarchy.
        /// </returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T) return (T)(BaseMap)this;
            if (this.content == null) return null;
            return this.content.FirstDescendentOfType<T>();
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap" /> in this <see cref="BaseMap" /><br />
        /// which has a specified key. This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap" /> that we wish to find the first instance of</typeparam>
        /// <param name="key">The key of the typed item that we require</param>
        /// <returns>
        /// The first instance of type <typeparamref name="T" /> found in the hierarchy.
        /// </returns>
        internal override T FirstDescendentOfType<T>(string key)
        {
            if (this is T && this.Key == key) return (T)(BaseMap)this;
            if (this.content == null) return null;
            return this.content.FirstDescendentOfType<T>(key);
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap" /> in this <see cref="StackPanel" /><br />.
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap" /> that we wish to find</typeparam>
        /// <param name="list">The list to be updated</param>
        internal override void AddDescendentsOfType<T>(ref List<T> list)
        {
            if (this is T)
            {
                list.Add((T)(BaseMap)this);
            }

            if (this.content != null)
            {
                this.content.AddDescendentsOfType<T>(ref list);
            }
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        /// <returns></returns>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Template");
        }

        #endregion Internal Methods
    }
}

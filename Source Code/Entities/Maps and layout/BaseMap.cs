namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Base class for Maps.
    /// Provides a style that can be applied to the map which dictates background colour and borders.
    /// </summary>
    public abstract class BaseMap : DataContextBase, IExcelPreparable, IResource
    {
        #region Private Fields

        private string key;
        private string mapId;
        private string partId;
        private string templateId;
        private object enabled;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets the context of the data for this instance (This is a property that can be used with binding)
        /// </summary>
        public object Enabled
        {
            get { return BindingContainer.EvaluateIfRequired(this.enabled, this.DataContext); }
            set { this.enabled = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a value that uniquely identifies the map - (CH perhaps?)
        /// </summary>
        [Obsolete("Check the use of MapId and similarities with Key")]
        public string MapId
        {
            get { return this.mapId; }
            set { this.mapId = value; }
        }

        /// <summary>
        /// Gets or sets the part id.
        /// When set it is used at runtime to set the DataContext of this instance from a DataPart with this Id
        /// </summary>
        /// <value>
        /// The part id.
        /// </value>
        public string PartId 
        {
            get { return this.partId; }
            set { this.partId = value; }
        }

        /// <summary>
        /// Gets or sets a value that identifies this element as a potential template, making it re-assignable to different data parts via metadata mapping.<br/>
        /// </summary>
        public string TemplateId
        {
            get { return this.templateId; }
            set { this.templateId = value; }
        }

        /// <summary>
        /// Gets or sets a string value that identifies this element as a potentially re-usable resource which can be looked up and implemented by Key<br/>
        /// </summary>
        public string Key
        {
            get { return this.key; }
            set { this.key = value; }
        }

        #endregion Public Properties

        #region Internal Properties

        private string resourceKey;

        /// <summary>
        /// Gets or sets the resource key.
        /// This identifies the key of the map at the root level of the resource and is used to return designer information such as the file name used by charts
        /// </summary>
        internal string ResourceKey
        {
            get { return this.resourceKey; }
            set { this.resourceKey = value; } 
        }

        /// <summary>
        /// Used internally at load time
        /// </summary>
        internal TemplateCollection TemplateCollection { get; set; }

        /// <summary>
        /// Gets a value which indicates if this <see cref="BaseMap"/> derived element is visual, i.e. will be written into Exel.
        /// </summary>
        internal virtual bool IsVisual
        {
            get { return true; }
        }

        /// <summary>
        /// Gets or sets a value which indicates whether this element has been prepared as part of data part preparation.
        /// </summary>
        internal bool Prepared { get; set; }

        #endregion Internal Properties

        #region Internal Methods

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy, or null if none found.</returns>
        internal abstract T FirstDescendentOfType<T>() where T : BaseMap;

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// which has a specified key. This includes this instance.
        /// </summary>
        /// <param name="key">The key of the typed item that we require</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy, or null if none found.</returns>
        internal abstract T FirstDescendentOfType<T>(string key) where T : BaseMap;

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>.
        /// This includes this instance.
        /// </summary>
        /// <param name="list">The list to be updated</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find</typeparam>
        internal abstract void AddDescendentsOfType<T>(ref List<T> list) where T : BaseMap;

        /// <summary>
        /// Builds a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>.
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find</typeparam>
        /// <returns>All instances of type <typeparamref name="T"/> found in the hierarchy, or null if none found.</returns>
        internal List<T> AllDescendentsOfType<T>() where T : BaseMap
        {
            var list = new List<T>();
            this.AddDescendentsOfType<T>(ref list);
            return list;
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal abstract string GetContainerType();

        /// <summary>
        /// Adds key to container type (if exists)
        /// </summary>
        /// <param name="containerType">Text representation of what the conainer is representing</param>
        /// <returns>Updated container type text</returns>
        internal string GetContainerTypeWithKey(string containerType)
        {
            return string.IsNullOrEmpty(this.key) ? containerType : string.Format("{0}[Key={1}]", containerType, this.key);
        }

        #endregion Internal Methods
    }
}

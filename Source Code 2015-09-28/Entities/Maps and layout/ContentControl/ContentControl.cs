namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Markup;

    /// <summary>
    /// Represents a placeholder for any <see cref="Content"/> based content.<br/>
    /// Used to create an instance of a Map from one defined in the Maps property of the <see cref="TemplateCollection"/><br/>
    /// Uses the Key property of the  for lookup of <see cref="Content"/>.
    /// </summary>
    [ContentProperty("Columns")]
    public class ContentControl : BaseMap
    {
        #region Private Fields

        private string contentKey;
        private BaseMap content;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Default ctor.
        /// </summary>
        public ContentControl()
            : base()
        {
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Key which is used to look up a <see cref="Content"/> derived element stored as a resource in the <see cref="TemplateCollection"/>.
        /// </summary>
        public string ContentKey
        {
            get { return this.contentKey; }
            set { this.contentKey = value; }
        }

        #endregion Public Properties

        #region Internal Methods

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="Content"/> in this <see cref="Content"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="Content"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        internal override T FirstDescendentOfType<T>()
        {
            if (this is T) return (T)(BaseMap)this;
            if (this.content == null) return null;
            return this.content.FirstDescendentOfType<T>();
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
            if (this.content == null) return null;
            return this.content.FirstDescendentOfType<T>(key);
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

            if (this.content != null)
            {
                this.content.AddDescendentsOfType<T>(ref list);
            }
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("ContentControl");
        }

        #endregion Internal Methods

        #region Internal Properties

        /// <summary>
        /// Get/set the map which results from the lookup
        /// </summary>
        public BaseMap Content
        {
            get { return this.content; }
            set { this.content = value; }
        }

        #endregion
    }
}

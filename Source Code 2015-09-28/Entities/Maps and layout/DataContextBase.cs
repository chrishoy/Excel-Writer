namespace ExcelWriter
{
    /// <summary>
    /// Base class for a class which has a DataContext.
    /// </summary>
    public abstract class DataContextBase
    {
        #region Private Fields

        private object dataContext;

        #endregion Private Fields

        #region Internal Properties

        /// <summary>
        /// Gets or sets the parent data context. Used for binding the DataContext itself.
        /// </summary>
        internal object ParentDataContext { get; set; }

        #endregion Internal Properties

        #region Public Properties

        /// <summary>
        /// Gets or sets the context of the data for this instance (This is a property that can be used with binding)
        /// </summary>
        public object DataContext
        {
            get { return BindingContainer.EvaluateIfRequired(this.dataContext, this.ParentDataContext); }
            set { this.dataContext = BindingContainer.CreateIfRequired(value); }
        }

        #endregion Public Properties
    }
}

namespace ExcelWriter
{
    using System;
    using System.Windows.Media;
    using System.Windows;

    /// <summary>
    /// Base class for map styles
    /// </summary>
    public abstract class StyleBase : ICloneable, IResource
    {
        #region Private Fields

        private string key;
        private string basedOnKey;
        private Color? backgroundColour;
        private Color? borderColour;
        private Thickness? borderThickness;

        // Bindable
        private object dataContext;

        #endregion Private Fields

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
        /// Get/set the background colour of this map.
        /// </summary>
        public Color? BackgroundColour
        {
            get { return this.backgroundColour; }
            set { this.backgroundColour = value; }
        }

        /// <summary>
        /// Get/set the key of the <see cref="Style"/> that this <see cref="Style"/> is based on.
        /// </summary>
        public string BasedOnKey
        {
            get { return this.basedOnKey; }
            set { this.basedOnKey = value; }
        }

        /// <summary>
        /// Get/set the colour of the border set around this map using BorderThickness
        /// </summary>
        public Color? BorderColour
        {
            get { return this.borderColour; }
            set 
            {
                this.borderColour = value;

            }
        }

        /// <summary>
        /// Get/set the thickness of each border around this map.
        /// </summary>
        public Thickness? BorderThickness
        {
            get { return this.borderThickness; }
            set { this.borderThickness = value; }
        }

        /// <summary>
        /// Get/set the key for this style within a dictionary.
        /// </summary>
        public string Key
        {
            get { return this.key; }
            set { this.key = value; }
        }

        /// <summary>
        /// Gets or sets the parent data context. Used for binding the DataContext itself.
        /// </summary>
        internal object ParentDataContext { get; set; }

        #endregion Public Properties

        #region IClonable members

        /// <summary>
        /// Create and return a copy of this instance.
        /// </summary>
        /// <returns></returns>
        public abstract object Clone();

        #endregion IClonable members
    }
}

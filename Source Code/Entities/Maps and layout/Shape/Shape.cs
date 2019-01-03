namespace ExcelWriter
{
    /// <summary>
    /// Represents a XAML definition for a shape that can be exported to an Excel worksheet.<br/>
    /// A <see cref="Shape"/> must be based on a <see cref="ShapeTemplate"/>.
    /// </summary>
    public class Shape : PositionableMap
    {
        #region Private Fields

        // Bindable
        private object shapeTemplateKey;
        private object fillColour;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets the name of the <see cref="ShapeTemplate"/> resource which will be used to derive this shape.<br/>
        /// This a Bindable property.
        /// </summary>
        public object ShapeTemplateKey
        {
            get { return BindingContainer.EvaluateIfRequired(this.shapeTemplateKey, this.DataContext); }
            set { this.shapeTemplateKey = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Gets or sets a fill colour for the shape.
        /// </summary>
        public object FillColour
        {
            get { return BindingContainer.EvaluateIfRequired(this.fillColour, this.DataContext); }
            set { this.fillColour = BindingContainer.CreateIfRequired(value); }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Shape");
        }

        #endregion Internal Methods
    }
}

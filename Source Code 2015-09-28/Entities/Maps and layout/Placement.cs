namespace ExcelWriter
{
    /// <summary>
    /// Represents the size and position information about a <see cref="PositionableMap"/>
    /// </summary>
    public sealed class Placement : DataContextBase
    {
        #region Private Fields

        private object height;
        private object width;
        private object verticalOffset;
        private object horizontalOffset;

        #endregion Private Fields

        /// <summary>
        ///  Gets or sets a value that determines the width that the element hosted by this <see cref="PositionableMap"/> should attempt to be when exported to Excel.<br/>
        ///  Null is 'as per template'.
        /// </summary>
        public object Width
        {
            get { return BindingContainer.EvaluateIfRequired(this.width, this.DataContext); }
            set { this.width = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        ///  Gets or sets a value that determines the height that the element hosted by this <see cref="PositionableMap"/> should attempt to be when exported to Excel.<br/>
        ///  Null is 'as per template'.
        /// </summary>
        public object Height
        {
            get { return BindingContainer.EvaluateIfRequired(this.height, this.DataContext); }
            set { this.height = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        ///  Gets or sets a value that determines the offset from the top that the element hosted by this <see cref="PositionableMap"/> should attempt to be when exported to Excel.<br/>
        ///  Null is 'as per template'.
        /// </summary>
        public object VerticalOffset
        {
            get { return BindingContainer.EvaluateIfRequired(this.verticalOffset, this.DataContext); }
            set { this.verticalOffset = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        ///  Gets or sets a value that determines the height that the element hosted by this <see cref="PositionableMap"/> should attempt to be when exported to Excel.<br/>
        ///  Null is 'as per template'.
        /// </summary>
        public object HorizontalOffset
        {
            get { return BindingContainer.EvaluateIfRequired(this.horizontalOffset, this.DataContext); }
            set { this.horizontalOffset = BindingContainer.CreateIfRequired(value); }
        }

    }
}

namespace ExcelWriter
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents a definition for a shape template resource which can be used to create shape instances within an Excel document.
    /// </summary>
    public class ShapeTemplate : BaseMap
    {
        #region Private Fields

        private string templateShapeName;
        private string templateSheetName;

        #endregion Private Fields


        #region Public Properties

        /// <summary>
        /// Gets or sets the name of the shape in the DesignerFile where this shape template resides.
        /// </summary>
        /// <value>
        /// The name of the template shape.
        /// </value>
        public string TemplateShapeName
        {
            get { return this.templateShapeName; }
            set { this.templateShapeName = value; }
        }

        /// <summary>
        /// Gets or sets the name of the sheet in the DesignerFile where this shape template resides.
        /// </summary>
        /// <value>
        /// The name of the template sheet.
        /// </value>
        public string TemplateSheetName
        {
            get { return this.templateSheetName; }
            set { this.templateSheetName = value; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets a value which indicates that this <see cref="ShapeTemplate"/> is not visual, i.e. is never written into Exel
        /// </summary>
        internal override bool IsVisual
        {
            get { return false; }
        }

        #endregion Internal Properties

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
            return null;
            //NB! For the moment, we will ignore anything lower.
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

            return null;

            // NB! For the moment, we will ignore anything lower.
        }

        /// <summary>
        /// Updates a list of all instances of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="TableData"/><br/>.
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

            // For the moment, don't go any lower
        }

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("ShapeTemplate");
        }

        #endregion Internal Methods
    }
}

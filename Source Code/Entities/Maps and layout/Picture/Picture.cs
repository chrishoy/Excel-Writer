namespace ExcelWriter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Markup;

    using OpenXml.Excel.Model;
    
    /// <summary>
    /// Represents a XAML definition for an picture that can be exported to an Excel worksheet.<br/>
    /// A <see cref="Picture"/> must be based on a <see cref="PictureTemplate"/>.
    /// </summary>
    public class Picture : PositionableMap
    {
        #region Private Fields

        // Bindable
        private object pictureTemplateKey;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets the name of the <see cref="PictureTemplate"/> resource which will be used to derive this picture.<br/>
        /// This a Bindable property.
        /// </summary>
        public object PictureTemplateKey
        {
            get { return BindingContainer.EvaluateIfRequired(this.pictureTemplateKey, this.DataContext); }
            set { this.pictureTemplateKey = BindingContainer.CreateIfRequired(value); }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Gets text which is used (mainly for debugging) to identify what the container represents.
        /// </summary>
        internal override string GetContainerType()
        {
            return base.GetContainerTypeWithKey("Picture");
        }

        #endregion Internal Methods
    }
}

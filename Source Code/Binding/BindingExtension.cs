namespace ExcelWriter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Windows.Markup;

    public class Binding : BindingExtension {  }

    /// <summary>
    /// XAML Markup Extension which, when used with <see cref="Gam.MM.Framework.Export.Map"/> derived entities, will use reflection to read certain property values. <br/>
    /// NB! The property being reflected much be coded to work with this <see cref="BindingExtension"/>.
    /// </summary>
    public class BindingExtension : MarkupExtension, IEnumerable
    {
        private string path;
        private string xpath;
        private string stringFormat;

        #region Construction

        /// <summary>
        /// Ctor. Requires a path parameter.
        /// </summary>
        /// <param name="type"></param>
        public BindingExtension(object path)
            : base()
        {
            // Set path to the property that we will attempt to read using reflection
            this.path = (string)path;
        }

        /// <summary>
        /// Default ctor.
        /// </summary>
        public BindingExtension() : base()
        {

        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Get/set the path set on this <see cref="BindingExtension"/>
        /// </summary>
        public string Path
        {
            get { return this.path; }
            set { this.path = value; }
        }

        /// <summary>
        /// Get/set the xpath set on this <see cref="BindingExtension"/>
        /// </summary>
        public string XPath
        {
            get { return this.xpath; }
            set { this.xpath = value; }
        }

        /// <summary>
        /// Get/set a string format to be used when the binding resolves to a string value.
        /// </summary>
        public string StringFormat
        {
            get { return this.stringFormat; }
            set { this.stringFormat = value; }
        }

        #endregion Public Properties

        /// <summary>
        /// Returns an object that is set as the value of the target property for this markup extension.
        /// </summary>
        /// <param name="serviceProvider">Object that can provide services for the markup extension.</param>
        /// <returns>The object value to set on the property where the extension is applied.</returns>
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return (object)this;
        }

        /// <summary>
        /// Returns a string representation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            return string.Format("{0}:Path='{1}'", base.ToString(), this.Path);
        }

        /// <summary>
        /// Bit of a bodge to allow the xaml parser to parse enumerables in bindings.
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            return new List<object>().GetEnumerator();
        }
    }
}

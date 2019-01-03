using System;
using System.IO;
using System.Windows;
using System.Windows.Markup;
using System.Xml;
using System.Collections.Generic;

namespace ExcelWriter
{
    /// <summary>
    /// Contains a collection of XAML defined templates, which are used for writing data into Excel worksheets.
    /// </summary>
    [ContentProperty("Templates")]
    public sealed class TemplateCollection
    {
        #region Private Fields

        private Templates templates;
        private StylesCollection styleResources;
        private MapCollection maps;
        private CellStyleSelectorCollection cellStyleSelectors;
        private string templateFileName;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Default constructor
        /// </summary>
        public TemplateCollection()
        {
            this.templates = new Templates();
            this.styleResources = new StylesCollection();
            this.maps = new MapCollection();
            this.cellStyleSelectors = new CellStyleSelectorCollection();
        }

        #endregion Construction

        #region Public Propeties

        /// <summary>
        /// The location of the template file which contains the implementation of the excel templates
        /// </summary>
        public string TemplateFileName
        {
            get { return this.templateFileName; }
            set { this.templateFileName = value; }
        }

        /// <summary>
        /// Get the style resources that have be specified for this template
        /// </summary>
        public StylesCollection StyleResources
        {
            get { return this.styleResources; }
            set { this.styleResources = value; }
        }

        /// <summary>
        /// The list of the templates this metadata knows about
        /// </summary>
        public Templates Templates
        {
            get { return this.templates; }
            set { this.templates = value; }
        }

        /// <summary>
        /// The list of the maps this metadata knows about
        /// </summary>
        public MapCollection Maps
        {
            get { return this.maps; }
        }

        /// <summary>
        /// The list of the cellStyleSelectors this metadata knows about
        /// </summary>
        public CellStyleSelectorCollection CellStyleSelectors
        {
            get { return this.cellStyleSelectors; }
        }
                
        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets a string representation of the XAML that was used to create this element.
        /// </summary>
        internal string XamlString { get; set; }

        #endregion Internal Properties

        #region Public Methods

        /// <summary>
        /// Creates an instance of the ExcelTemplateCollection from a xaml string
        /// </summary>
        /// <param name="value">A xaml string representation of the ExcelTemplateCollection</param>
        /// <returns>An instant of ExcelTemplateCollection</returns>
        public static TemplateCollection Deserialize(string value)
        {
            using (var sr = new StringReader(value))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (TemplateCollection)XamlReader.Load(xr);
                }
            }
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Builds a list of elements of a specified type derived from <see cref="BaseMap"/> in this <see cref="TemplateCollection"/><br/>.
        /// This includes this instance.
        /// </summary>
        /// <param name="list">The list to be updated</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find</typeparam>
        internal IEnumerable<T> GetElementsOfType<T>() where T : BaseMap
        {
            var list = new List<T>();

            foreach (var element in this.Maps)
            {
                if (element is T)
                {
                    list.Add((T)(BaseMap)element);
                }
            }

            return list;

        }

        #endregion Internal Methods

    }
}

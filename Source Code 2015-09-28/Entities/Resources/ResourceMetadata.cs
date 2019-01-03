namespace ExcelWriter
{
    using System.IO;
    using System.Windows.Markup;
    using System.Xml;

    /// <summary>
    /// Contains a collection of XAML defined resources, which are used for writing data into documents.
    /// </summary>
    [ContentProperty("Resources")]
    public sealed class ResourceMetadata : IResourceContainer
    {
        #region Construction

        /// <summary>
        /// Default ctor.
        /// </summary>
        public ResourceMetadata()
        {
            this.Resources = new ResourceCollection();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the key which identifies a resource.
        /// </summary>
        public string Source { get; set; }

        /// <summary>
        /// The location of the template file which contains the implementation of the excel templates
        /// </summary>
        public string DesignerFileName { get; set; }

        /// <summary>
        /// The collection of resources supplied by this dictionary
        /// </summary>
        /// <value>
        /// The resources.
        /// </value>
        public ResourceCollection Resources { get; set; }
   
        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Creates an instance of the ExcelTemplateCollection from a xaml string
        /// </summary>
        /// <param name="value">A xaml string representation of the ExcelTemplateCollection</param>
        /// <returns>An instant of ExcelTemplateCollection</returns>
        public static ResourceMetadata Deserialize(string value)
        {
            using (var sr = new StringReader(value))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (ResourceMetadata)XamlReader.Load(xr);
                }
            }
        }

        #endregion Public Methods       
    }
}

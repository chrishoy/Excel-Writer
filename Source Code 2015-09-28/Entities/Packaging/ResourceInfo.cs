using System;

namespace ExcelWriter
{
    /// <summary>
    /// Stores all the information used to create an instance of BaseMap or a Template
    /// </summary>
    internal sealed class ResourceInfo
    {
        public ResourceInfo(string key, string resourceString, TemplateCollection templateCollection, string templateCollectionUri)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }
            if (string.IsNullOrEmpty(resourceString))
            {
                throw new ArgumentNullException("resourceString");
            }
            if (templateCollection == null)
            {
                throw new ArgumentNullException("templateCollection");
            }
            if (string.IsNullOrEmpty(templateCollectionUri))
            {
                throw new ArgumentNullException("templateCollectionUri");
            }

            this.Key = key;
            this.ResourceString = resourceString;
            this.TemplateCollection = templateCollection;
            this.TemplateCollectionUri = templateCollectionUri;
        }

        /// <summary>
        /// Gets the key that uniquely identifies the BaseMap or Template.
        /// </summary>
        internal string Key { get; private set; }

        /// <summary>
        /// Gets the resource string.
        /// </summary>
        internal string ResourceString { get; private set; }

        /// <summary>
        /// Gets the template collection.
        /// </summary>
        internal TemplateCollection TemplateCollection { get; private set; }

        /// <summary>
        /// Gets the name of the template file.
        /// </summary>
        internal string TemplateFileName { get { return this.TemplateCollection.TemplateFileName; } }

        /// <summary>
        /// Gets the resource string.
        /// </summary>
        internal string TemplateCollectionUri { get; private set; }
    }
}
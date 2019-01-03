namespace ExcelWriter
{
    using System;

    /// <summary>
    /// Stores all the information used to create an instance of BaseMap or a Template
    /// </summary>
    public sealed class ResourceData
    {
        public ResourceData(string key, string resourceType, IResource instance, string uri)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }
            if (instance == null) 
            {
                throw new ArgumentNullException("instance");
            }
            if (string.IsNullOrEmpty(resourceType))
            {
                throw new ArgumentNullException("resourceType");
            }

            this.Key = key;
            this.Instance = instance;
            this.IsInstance = true;
            this.ResourceType = resourceType;
            this.Uri = uri ?? "unknown";
        }

        public ResourceData(string key, string resourceType, string resourceString, string designerFileName, string uri)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }
            if (string.IsNullOrEmpty(resourceString))
            {
                throw new ArgumentNullException("resourceString");
            }
            if (string.IsNullOrEmpty(resourceType))
            {
                throw new ArgumentNullException("resourceType");
            }

            this.Key = key;
            this.DesignerFileName = designerFileName;
            this.IsInstance = false;
            this.ResourceString = resourceString;
            this.ResourceType = resourceType;
            this.Uri = uri ?? "unknown";
        }

        /// <summary>
        /// Gets the key that uniquely identifies the BaseMap or Template.
        /// </summary>
        internal string Key { get; private set; }

        internal string ResourceType { get; private set; }

        /// <summary>
        /// Gets the resource string.
        /// </summary>
        internal string ResourceString { get; private set; }

        /// <summary>
        /// Gets the name of the template file.
        /// </summary>
        internal string DesignerFileName { get; private set; }

        /// <summary>
        /// Gets the resource string.
        /// </summary>
        internal string Uri { get; private set; }

        internal IResource Instance { get; private set; }

        internal bool IsInstance { get; private set; }
    }
}
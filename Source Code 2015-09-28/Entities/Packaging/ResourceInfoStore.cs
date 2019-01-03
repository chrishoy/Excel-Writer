namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using System.Windows.Markup;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Stores resource info for Templates or BaseMaps
    /// </summary>
    internal sealed class ResourceInfoStore
    {
        private bool isMapStore;
        private Dictionary<string, ResourceInfo> resourceInfoDictionary;
        internal Dictionary<string, ExcelTemplateFileInfo> templateDocumentDictionary;

        public ResourceInfoStore(bool isMapStore)
        {
            this.isMapStore = isMapStore;
            this.resourceInfoDictionary = new Dictionary<string, ResourceInfo>();
            this.templateDocumentDictionary = new Dictionary<string, ExcelTemplateFileInfo>();
        }

        /// <summary>
        /// Add a new set of resource information to the store.
        /// </summary>
        /// <param name="key">The unique id of the resource</param>
        /// <param name="resourceString">A string of Xaml. Will be used to inflate on demand.</param>
        /// <param name="templateCollection">The parent TemplateCollection stores Style etc</param>
        /// <param name="templateCollectionUri">Handy to know the TemplateCollection Uri if there are duplicates</param>
        public void Add(string key, string resourceString, TemplateCollection templateCollection, string templateCollectionUri) 
        {
            if (resourceInfoDictionary.ContainsKey(key)) 
            {
                // handy to log where the existing is, so get the xaml uri
                throw new MetadataException(string.Format("Loading <{0}> - already loaded Map {1} in {2}", templateCollectionUri, key, resourceInfoDictionary[key].TemplateCollectionUri));
            }

            var resourceInfo = new ResourceInfo(key, resourceString, templateCollection, templateCollectionUri);
            this.resourceInfoDictionary.Add(key, resourceInfo);
        }

        public void AddTemplateFileData(string templateFileName, byte[] data)
        {
            // exception - cant add template file name from a map store
            if (this.isMapStore)
            {
                throw new MetadataException("Unable to add Template File to a Map Store");
            }

            // exception - already loaded
            if (this.templateDocumentDictionary.ContainsKey(templateFileName))
            {
                throw new MetadataException(string.Format("Already loaded template file {0}", templateFileName));
            }

            this.templateDocumentDictionary.Add(templateFileName, new ExcelTemplateFileInfo(templateFileName, data));
        }

        /// <summary>
        /// Try and return a new Template from the store for the supplied key
        /// </summary>
        /// <param name="key">The key find a Template for</param>
        /// <returns>A new instance of the Template</returns>
        public Template GetTemplateByKey(string key) 
        {
            // exception - cant return a template from a map store
            if (this.isMapStore) 
            {
                throw new MetadataException("Unable to return a Template from a Map Store");
            }

            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }

            // exception - no matche 
            if (!this.resourceInfoDictionary.ContainsKey(key))
            {
                throw new MetadataException(string.Format("No Template found in Key <{0}>", key));
            }

            Template template = null;

            using (var sr = new StringReader(this.resourceInfoDictionary[key].ResourceString))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    template = (Template)XamlReader.Load(xr);
                }
            }

            var templatesCollection = this.resourceInfoDictionary[key].TemplateCollection;

            template.CellStyleSelectors = templatesCollection.CellStyleSelectors;
            template.Maps = templatesCollection.Maps;
            template.MapStyles = templatesCollection.StyleResources;
            template.TemplateCollection = templatesCollection;

            return template;
        }

        /// <summary>
        /// Gets the template spreadsheet document by key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        public SpreadsheetDocument GetTemplateSpreadsheetDocumentByKey(string key)
        {
            string fileName = this.GetTemplateFileNameByKey(key);

            if (!this.templateDocumentDictionary.ContainsKey(fileName))
            {
                throw new MetadataException(string.Format("Template file {0} not found for TemplateId <{1}>. Please check the templates directory", fileName, key));
            }

            return this.templateDocumentDictionary[fileName].SpreadsheetDocument;
        }

        /// <summary>
        /// Try and return a new Template from the store for the supplied key
        /// </summary>
        /// <param name="key">The key find a Template for</param>
        /// <returns>A new instance of the Template</returns>
        public string GetTemplateFileNameByKey(string key)
        {
            // exception - cant return a template from a map store
            if (this.isMapStore)
            {
                throw new MetadataException("Unable to return a Template from a Map Store");
            }

            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }

            // exception - no matche 
            if (!this.resourceInfoDictionary.ContainsKey(key))
            {
                throw new MetadataException(string.Format("No Template found in Key <{0}>", key));
            }

            return this.resourceInfoDictionary[key].TemplateFileName;
        }

        /// <summary>
        /// Try and return a new BaseMap from the store for the supplied key
        /// </summary>
        /// <param name="key">The key find a BaseMap for</param>
        /// <returns>A new instance of the BaseMap</returns>
        public BaseMap GetMapByKey(string key)
        {
            // exception - cant return a map from a template store
            if (!this.isMapStore)
            {
                throw new MetadataException("Unable to return a Map from a Template Store");
            }

            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }

            // exception - no matche 
            if (!this.resourceInfoDictionary.ContainsKey(key))
            {
                throw new MetadataException(string.Format("No Map found in Key <{0}>", key));
            }

            using (var sr = new StringReader(this.resourceInfoDictionary[key].ResourceString))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (BaseMap)XamlReader.Load(xr);
                }
            }
        }

        public void Validate()
        {
            if (!this.isMapStore) 
            {
                foreach (var item in this.resourceInfoDictionary.Values)
                {
                    Output(string.Format("Checking TemplateId <{0}>", item.Key));

                    if (!this.templateDocumentDictionary.ContainsKey(item.TemplateFileName))
                    {
                        throw new MetadataException(string.Format("TemplateFileName <{0}> cannot be found", item.TemplateFileName));
                    }

                    Output(string.Format("TemplateFileName <{0}> found", item.TemplateFileName));
                }
            }
        }

        public void Flush()
        {
            foreach (var doc in this.templateDocumentDictionary.Values)
            {
                doc.Flush();
            }
        }

        private static void Output(string message)
        {            
            message = string.Format("{0} : {1}", DateTime.Now.ToString("HH:mm:ss FFFF"), message);
            Console.WriteLine(message);
        }
    }
}

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Windows.Markup;
    using System.Xml;
    using System.Xml.Linq;

    using DocumentFormat.OpenXml.Packaging;

    using Constants;
    using OpenXml.Excel;

    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
    using ExcelModel = OpenXml.Excel.Model;

    /// <summary>
    /// Stores resource info for Templates or BaseMaps
    /// </summary>
    public sealed class ResourceStore
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ResourceStore"/> class.
        /// </summary>
        public ResourceStore()
        {
            this.ResourceDataDictionary = new Dictionary<string, ResourceData>();
            this.OpenXmlPackageDictionary = new Dictionary<string, OpenXmlPackageInfo>();
            this.ChartModelDictionary = new Dictionary<string, ExcelModel.ChartModel>();
            this.ShapeModelDictionary = new Dictionary<string, ExcelModel.ShapeModel>();
            this.PictureModelDictionary = new Dictionary<string, ExcelModel.PictureModel>();
        }

        /// <summary>
        /// Gets the resource data dictionary.
        /// </summary>
        /// <value>
        /// The resource data dictionary.
        /// </value>
        internal Dictionary<string, ResourceData> ResourceDataDictionary { get; private set; }
        /// <summary>
        /// Gets the open XML package dictionary.
        /// </summary>
        /// <value>
        /// The open XML package dictionary.
        /// </value>
        internal Dictionary<string, OpenXmlPackageInfo> OpenXmlPackageDictionary { get; private set; }

        /// <summary>
        /// Gets the chart model dictionary.
        /// </summary>
        /// <value>
        /// The chart model dictionary.
        /// </value>
        internal Dictionary<string, ExcelModel.ChartModel> ChartModelDictionary { get; private set; }

        /// <summary>
        /// Gets the shape model dictionary.
        /// </summary>
        /// <value>
        /// The shape model dictionary.
        /// </value>
        internal Dictionary<string, ExcelModel.ShapeModel> ShapeModelDictionary { get; private set; }

        /// <summary>
        /// Gets the picture model dictionary.
        /// </summary>
        /// <value>
        /// The picture model dictionary.
        /// </value>
        internal Dictionary<string, ExcelModel.PictureModel> PictureModelDictionary { get; private set; }

        /// <summary>
        /// Creates a fully populated <see cref="ResourceStore" /> which contains all resources defined in a <see cref="TemplateCollection" />.
        /// This is for backward compatibility during processing of a <see cref="TemplateCollection" />.
        /// </summary>
        /// <param name="templateCollection">The template collection.</param>
        /// <param name="templatePackage">The template package.</param>
        /// <returns>
        /// A fully populated <see cref="ResourceStore" />
        /// </returns>
        internal static ResourceStore Create(TemplateCollection templateCollection, ExcelTemplatePackage templatePackage)
        {
            if (templateCollection == null || string.IsNullOrEmpty(templateCollection.XamlString))
            {
                return null;
            }

            ResourceStore resourceStore = new ResourceStore();
            var document = XDocument.Parse(templateCollection.XamlString);

            // Add in the defined Styles
            foreach (var resource in templateCollection.StyleResources)
            {
                resourceStore.Add(resource.Key, ResourceTypeNames.StyleBase, resource, null);
            }

            // Add in CellStyleSelectors
            foreach (var resource in templateCollection.CellStyleSelectors)
            {
                resourceStore.Add(resource.Key, ResourceTypeNames.CellStyleSelector, resource, null);
            }

            // Loop over each all elements, adding any that are defined in the Maps collection or identified as being Template resources
            foreach (var element in document.Root.Elements())
            {
                if (element.NodeType != XmlNodeType.Element || element.Name == null)
                {
                    continue;
                }

                if (element.Name.LocalName.CompareTo("TemplateCollection.Maps") == 0)
                {
                    foreach (var mapElement in element.Elements())
                    {
                        AddElementToStore(templateCollection, resourceStore, mapElement);
                    }
                }
                else if (element.Name.LocalName.CompareTo("Template") == 0)
                {
                    var id = element.Attribute("TemplateId").Value;
                    resourceStore.Add(id, element.Name.LocalName, element.ToString(), templateCollection.TemplateFileName, null);
                }
            }

            // Add the TemplateFile as a designer file, which can be used by maps as a resource
            if (!string.IsNullOrEmpty(templateCollection.TemplateFileName))
            {
                var data = templatePackage.TemplateResourceStore.templateDocumentDictionary[templateCollection.TemplateFileName].Data;
                resourceStore.AddDesignerFileData(templateCollection.TemplateFileName, data);
            }

            // Populate chart and shape models from the supplied resource designer files and defined templates,
            // such as ChartTemplate and ShapeTempaltes.
            resourceStore.PopulateModels();

            return resourceStore;
        }

        /// <summary>
        /// Adds the element to store.
        /// </summary>
        /// <param name="templateCollection">The template collection.</param>
        /// <param name="resourceStore">The resource store.</param>
        /// <param name="element">The element.</param>
        private static void AddElementToStore(TemplateCollection templateCollection, ResourceStore resourceStore, XElement element)
        {
            var keyAttrib = element.Attribute("Key");
            if (keyAttrib != null && !string.IsNullOrEmpty(keyAttrib.Value))
            {
                resourceStore.Add(keyAttrib.Value, element.Name.LocalName, element.ToString(), templateCollection.TemplateFileName, null);
            }
        }

        /// <summary>
        /// Parses the specified resource string and creates a resource store from it
        /// </summary>
        /// <param name="resourceString">The resource string.</param>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        /// <exception cref="MetadataException"></exception>
        internal static ResourceStore Parse(string resourceString, string uri)
        {
            ResourceStore resourceStore = new ResourceStore();

            // Deserialise the template xaml
            IResourceContainer resourceContainer = null;

            using (var sr = new StringReader(resourceString))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    resourceContainer = (IResourceContainer)XamlReader.Load(xr);
                }
            }

            if (resourceContainer == null)
            {
                return resourceStore;
            }

            var indecesOfElementsToAdd = ExtractCellStyleSelectors(uri, resourceStore, resourceContainer);

            int idx = 0;
            var document = XDocument.Parse(resourceString);

            XElement rootElement = null;
            if (resourceContainer is ResourceMetadata)
            {
                rootElement = document.Root;
            }
            else if (resourceContainer is ExcelDocumentMetadata)
            {
                // initially check for a ResourceCollection node
                // there will be one of these if there is a MergeResources going on
                rootElement = (from d in document.Descendants()
                               where d.Name.LocalName.CompareTo("ResourceCollection") == 0
                               select d).FirstOrDefault();

                if (rootElement == null)
                {
                    // if no merge resources then look for a ExcelDocumentMetadata.Resources node
                    // all style etc will be located directly beneath this
                    rootElement = (from d in document.Descendants()
                                   where d.Name.LocalName.CompareTo("ExcelDocumentMetadata.Resources") == 0
                                   select d).FirstOrDefault();
                }
            }
            else
            {
                throw new MetadataException(string.Format("Unexpected type in ResourceStore.Parse <{0}>", resourceContainer.GetType().ToString()));
            }

            if (rootElement != null)
            {
                foreach (var element in rootElement.Elements())
                {
                    // only store those elements that we havent stored an instance of above
                    // basically everything other that Styles and CellStyleSelectors
                    if (indecesOfElementsToAdd.Contains(idx))
                    {
                        var keyAttrib = element.Attribute("Key");
                        if (keyAttrib != null && !string.IsNullOrEmpty(keyAttrib.Value))
                        {
                            // TODO: Optimisation - DesignerFileName could reference an entry in a string dictionary
                            resourceStore.Add(keyAttrib.Value, element.Name.LocalName, element.ToString(), resourceContainer.DesignerFileName, uri);
                        }
                    }
                    idx++;
                }
            }

            return resourceStore;
        }

        /// <summary>
        /// Builds and returns a list of indexes (into the resource store) of elements to add be added.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <param name="resourceStore">The resource store.</param>
        /// <param name="rd">The rd.</param>
        /// <returns></returns>
        private static List<int> ExtractCellStyleSelectors(string uri, ResourceStore resourceStore, IResourceContainer rd)
        {
            var indecesOfElementsToAdd = new List<int>();

            int index = 0;
            foreach (var resource in rd.Resources)
            {
                // CellStyleSelectors allow us to 'look up' cell styles depending on the underlying data
                // For this reason, they are isolated and added using the ResourceTypeName of CellStyleSelector
                if (resource is CellStyleSelector)
                {
                    resourceStore.Add(resource.Key, ResourceTypeNames.CellStyleSelector, resource, uri);
                }
                //else if (resource is StyleBase)
                //{
                //    resourceStore.Add(resource.Key, ResourceTypeNames.StyleBase, resource, uri);
                //}
                else
                {
                    indecesOfElementsToAdd.Add(index);
                }
                index++;
            }
            return indecesOfElementsToAdd;
        }

        /// <summary>
        /// Merges the specified resource store with this instance.
        /// </summary>
        /// <param name="resourceStore">The resource store.</param>
        internal void Merge(ResourceStore resourceStore)
        {
            foreach (var rd in resourceStore.ResourceDataDictionary.Values)
            {
                this.Add(rd);
            }

            foreach (var packageInfo in resourceStore.OpenXmlPackageDictionary.Values)
            {
                this.AddDesignerFileData(packageInfo.FileName, packageInfo.Data);
            }
        }

        /// <summary>
        /// Adds the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="resourceType">Type of the resource.</param>
        /// <param name="instance">The instance.</param>
        /// <param name="uri">The URI.</param>
        internal void Add(string key, string resourceType, IResource instance, string uri)
        {
            var rd = new ResourceData(key, resourceType, instance, uri);
            this.Add(rd);
        }

        /// <summary>
        /// Add a new set of resource information to the store.
        /// </summary>
        /// <param name="key">The unique id of the resource</param>
        /// <param name="elementName">Name of the element.</param>
        /// <param name="resourceString">A string of Xaml. Will be used to inflate on demand.</param>
        /// <param name="designerFileName">Name of the designer file.</param>
        /// <param name="uri">The URI.</param>
        internal void Add(string key, string elementName, string resourceString, string designerFileName, string uri)
        {
            var rd = new ResourceData(key, elementName, resourceString, designerFileName, uri);
            this.Add(rd);
        }

        /// <summary>
        /// Adds the specified resource data.
        /// </summary>
        /// <param name="resourceData">The resource data.</param>
        /// <exception cref="ArgumentNullException">resourceData</exception>
        /// <exception cref="MetadataException"></exception>
        internal void Add(ResourceData resourceData)
        {
            if (resourceData == null)
            {
                throw new ArgumentNullException("resourceData");
            }

            if (this.ResourceDataDictionary.ContainsKey(resourceData.Key))
            {
                // handy to log where the existing is, so get the xaml uri
                throw new MetadataException(string.Format("Loading <{0}> - already loaded Map {1} in {2}", resourceData.Uri, resourceData.Key, this.ResourceDataDictionary[resourceData.Key].Uri));
            }

            this.ResourceDataDictionary.Add(resourceData.Key, resourceData);
        }

        /// <summary>
        /// Adds the designer file data.
        /// </summary>
        /// <param name="designerFileName">Name of the designer file.</param>
        /// <param name="data">The data.</param>
        /// <exception cref="MetadataException"></exception>
        internal void AddDesignerFileData(string designerFileName, byte[] data)
        {
            // exception - already loaded
            if (this.OpenXmlPackageDictionary.ContainsKey(designerFileName))
            {
                throw new MetadataException(string.Format("Already loaded template file {0}", designerFileName));
            }

            this.OpenXmlPackageDictionary.Add(designerFileName, new OpenXmlPackageInfo(designerFileName, data));
        }

        /// <summary>
        /// Gets the map styles.
        /// </summary>
        /// <returns></returns>
        internal IEnumerable<StyleBase> GetMapStyles()
        {
            var mapStyles = from rd in this.ResourceDataDictionary.Values
                            where rd.ResourceType.CompareTo(ResourceTypeNames.StyleBase) == 0
                            || rd.ResourceType.CompareTo(ResourceTypeNames.Style) == 0
                            || rd.ResourceType.CompareTo(ResourceTypeNames.CellStyle) == 0
                            select rd;

            foreach (var rd in mapStyles)
            {
                if (rd.IsInstance)
                {
                    yield return (StyleBase)rd.Instance;
                }
                else
                {
                    yield return GetResourceByData<StyleBase>(rd);
                }
            }
        }

        /// <summary>
        /// Try and return a new Resource from the store for the supplied key
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key">The key to find a resource for</param>
        /// <returns>
        /// A new instance of the resource
        /// </returns>
        /// <exception cref="ArgumentNullException">key</exception>
        /// <exception cref="MetadataException"></exception>
        internal T GetResourceByKey<T>(string key) where T : IResource
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }

            // exception - no matche 
            if (!this.ResourceDataDictionary.ContainsKey(key))
            {
                throw new MetadataException(string.Format("No Resource found for Key <{0}>", key));
            }

            var rd = this.ResourceDataDictionary[key];

            if (rd.IsInstance)
            {
                return (T)rd.Instance;
            }
            else
            {
                return GetResourceByData<T>(rd);
            }
        }

        /// <summary>
        /// Gets the resource by data.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rd">The rd.</param>
        /// <returns></returns>
        private static T GetResourceByData<T>(ResourceData rd) where T : IResource
        {
            using (var sr = new StringReader(rd.ResourceString))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (T)XamlReader.Load(xr);
                }
            }
        }

        /// <summary>
        /// Gets the template spreadsheet document by key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        /// <exception cref="MetadataException"></exception>
        internal SpreadsheetDocument GetDesignerSpreadsheetDocumentByKey(string key)
        {
            string fileName = this.GetDesignerFileNameByKey(key);

            if (!this.OpenXmlPackageDictionary.ContainsKey(fileName))
            {
                throw new MetadataException(string.Format("Template file {0} not found for TemplateId <{1}>. Please check the templates directory", fileName, key));
            }

            return this.OpenXmlPackageDictionary[fileName].Package;
        }

        /// <summary>
        /// Try and return a new Template from the store for the supplied key
        /// </summary>
        /// <param name="key">The key find a Template for</param>
        /// <returns>
        /// A new instance of the Template
        /// </returns>
        /// <exception cref="ArgumentNullException">key</exception>
        /// <exception cref="MetadataException"></exception>
        internal string GetDesignerFileNameByKey(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentNullException("key");
            }

            // exception - no match
            if (!this.ResourceDataDictionary.ContainsKey(key))
            {
                throw new MetadataException(string.Format("No Template found in Key <{0}>", key));
            }

            return this.ResourceDataDictionary[key].DesignerFileName;
        }

        /// <summary>
        /// Resolves the references.
        /// </summary>
        /// <param name="resourceMetadataList">List of all the <see cref="ResourceMetadata" /> from which resources will be extracted into this <see cref="ResourceStore" /></param>
        /// <param name="resourcePackage"><see cref="ResourcePackage" /> which holds keyed xaml resources</param>
        internal void ResolveReferences(IEnumerable<ResourceMetadata> resourceMetadataList, ResourcePackage resourcePackage)
        {
            this.ExtractResources(resourceMetadataList, resourcePackage);
            this.PopulateModels();
        }

        /// <summary>
        /// Extracts the resources.
        /// </summary>
        /// <param name="resourceMetadataList">List of all the <see cref="ResourceMetadata" /> from which resources will be extracted into this <see cref="ResourceStore" /></param>
        /// <param name="resourcePackage"><see cref="ResourcePackage" /> which holds keyed xaml resources</param>
        private void ExtractResources(IEnumerable<ResourceMetadata> resourceMetadataList, ResourcePackage resourcePackage)
        {
            // Go through list of Resource files specified in the metadata package
            foreach (var rm in resourceMetadataList)
            {
                if (string.IsNullOrEmpty(rm.Source))
                {
                    continue;
                }

                // Get Resources from the supplied package
                var searchUri = string.Format("/Metadata/{0}.xaml", rm.Source);
                var matches = GetResourceDataForUri(resourcePackage, searchUri);

                // Build a list of designer files specified in those resource files
                List<string> designerFiles = new List<string>();
                foreach (var match in matches)
                {
                    // Add the matching resource to this store.
                    this.Add(match);

                    // If the resource has a DesignerFile, then add to a list, so we can load data about those designer files.
                    if (!string.IsNullOrEmpty(match.DesignerFileName) && !designerFiles.Contains(match.DesignerFileName))
                    {
                        designerFiles.Add(match.DesignerFileName);
                    }
                }

                // Go through the list of designer files, adding the designer file
                foreach (var fileName in designerFiles)
                {
                    if (resourcePackage.ResourceStore.OpenXmlPackageDictionary.ContainsKey(fileName) &&
                        this.OpenXmlPackageDictionary.ContainsKey(fileName) == false)
                    {
                        this.AddDesignerFileData(fileName, resourcePackage.ResourceStore.OpenXmlPackageDictionary[fileName].Data);
                    }
                }
            }
        }

        /// <summary>
        /// Populates any models from which Excel elements will be generated.
        /// TemplateFileName required for compatibility with legacy TempalteCollectio derrived maps.
        /// </summary>
        private void PopulateModels()
        {
            foreach (var item in this.ResourceDataDictionary)
            {
                ResourceData resourceData = item.Value;

                // Process ChartTemplate elements
                if (resourceData.ResourceType == "ChartTemplate")
                {
                    ChartTemplate chartTemplate = this.GetResourceByKey<ChartTemplate>(resourceData.Key);

                    if (chartTemplate != null)
                    {
                        WorksheetPart designerWorksheetPart = null;

                        // Get the worksheet on which the ChartTemplate resides (Alt. look at ExcelSheetMapper.GetDesignerWorksheet...)
                        var designerSpreadsheetDocument = GetDesignerSpreadsheetDocumentByKey(chartTemplate.Key);
                        if (designerSpreadsheetDocument != null)
                        {
                            designerWorksheetPart = designerSpreadsheetDocument.GetWorksheetPart(chartTemplate.TemplateSheetName);
                            if (designerWorksheetPart != null)
                            {
                                ExcelModel.ChartModel chartModel = ExcelModel.ChartModel.GetChartModel(designerWorksheetPart.Worksheet, chartTemplate.TemplateChartName);
                                if (chartModel != null)
                                {
                                    this.ChartModelDictionary.Add(item.Key, chartModel);
                                }
                            }
                        }
                    }
                }

                // Process ShapeTemplate elements
                if (resourceData.ResourceType == "ShapeTemplate")
                {
                    ShapeTemplate shapeTemplate = this.GetResourceByKey<ShapeTemplate>(resourceData.Key);

                    if (shapeTemplate != null)
                    {
                        WorksheetPart designerWorksheetPart = null;

                        // Get the worksheet on which the ShapeTemplate resides (Alt. look at ExcelSheetMapper.GetDesignerWorksheet...)
                        var designerSpreadsheetDocument = GetDesignerSpreadsheetDocumentByKey(shapeTemplate.Key);
                        if (designerSpreadsheetDocument != null)
                        {
                            designerWorksheetPart = designerSpreadsheetDocument.GetWorksheetPart(shapeTemplate.TemplateSheetName);
                            if (designerWorksheetPart != null)
                            {
                                ExcelModel.ShapeModel shapeModel = ExcelModel.ShapeModel.GetShapeModel(designerWorksheetPart.Worksheet, shapeTemplate.TemplateShapeName);
                                if (shapeModel != null)
                                {
                                    this.ShapeModelDictionary.Add(item.Key, shapeModel);
                                }
                            }
                        }
                    }
                }

                // Process PictureTemplate elements
                if (resourceData.ResourceType == "PictureTemplate")
                {
                    PictureTemplate pictureTemplate = this.GetResourceByKey<PictureTemplate>(resourceData.Key);

                    if (pictureTemplate != null)
                    {
                        WorksheetPart designerWorksheetPart = null;

                        // Get the worksheet on which the ShapeTemplate resides (Alt. look at ExcelSheetMapper.GetDesignerWorksheet...)
                        var designerSpreadsheetDocument = GetDesignerSpreadsheetDocumentByKey(pictureTemplate.Key);
                        if (designerSpreadsheetDocument != null)
                        {
                            designerWorksheetPart = designerSpreadsheetDocument.GetWorksheetPart(pictureTemplate.TemplateSheetName);
                            if (designerWorksheetPart != null)
                            {
                                ExcelModel.PictureModel pictureModel = ExcelModel.PictureModel.GetPictureModel(designerWorksheetPart.Worksheet, pictureTemplate.TemplatePictureName);
                                if (pictureModel != null)
                                {
                                    this.PictureModelDictionary.Add(item.Key, pictureModel);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets resource data contained in a <see cref="ResourcePackage" />  which has a specified uri.
        /// </summary>
        /// <param name="resourcePackage">The resource package.</param>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        private static IEnumerable<ResourceData> GetResourceDataForUri(ResourcePackage resourcePackage, string uri)
        {
            var matches = from rd in resourcePackage.ResourceStore.ResourceDataDictionary.Values
                          where rd.Uri.CompareTo(uri) == 0
                          select rd;
            return matches;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <exception cref="MetadataException"></exception>
        internal void Validate()
        {
            foreach (var item in this.ResourceDataDictionary.Values)
            {
                Output(string.Format("Checking TemplateId <{0}>", item.Key));
                if (!string.IsNullOrEmpty(item.DesignerFileName))
                {
                    if (!this.OpenXmlPackageDictionary.ContainsKey(item.DesignerFileName))
                    {
                        throw new MetadataException(string.Format("DesignerFileName <{0}> cannot be found", item.DesignerFileName));
                    }
                }

                Output(string.Format("TemplateFileName <{0}> found", item.DesignerFileName));
            }
        }

        /// <summary>
        /// Flushes this instance.
        /// </summary>
        internal void Flush()
        {
            foreach (var doc in this.OpenXmlPackageDictionary.Values)
            {
                doc.Flush();
            }
        }

        /// <summary>
        /// Outputs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        private static void Output(string message)
        {
            message = string.Format("{0} : {1}", DateTime.Now.ToString("HH:mm:ss FFFF"), message);
            Console.WriteLine(message);
        }
    }
}

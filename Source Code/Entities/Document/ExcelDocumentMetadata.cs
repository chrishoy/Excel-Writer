namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.IO;
    using System.Xml;
    using System.Windows.Markup;
    using System.Xml.Linq;

    /// <summary>
    /// 
    /// </summary>
    public class ExcelDocumentMetadata : DocumentMetadataBase, IResourceContainer
    {
        private Dictionary<string, string> sheetResourceStore;

        public ExcelDocumentMetadata()
        {
            this.Resources = new ResourceCollection();

            this.Sheets = new SheetCollection();
            this.sheetResourceStore = new Dictionary<string, string>();
        }

        public string DesignerFileName { get; set; }

        public ResourceCollection Resources { get; set; }

        public ResourceStore ResourceStore { get; internal set; }

        public SheetCollection Sheets { get; set; }

        public void MergeResources(ResourcePackage package)
        {
            this.ResourceStore.ResolveReferences(this.Resources.MergeResources, package);

        }

        /// <summary>
        /// Build a distinct list of part ids by walking the down each sheet's tree
        /// </summary>
        /// <returns></returns>
        public override IEnumerable<string> GetPartIds()
        {
            var result = new List<string>();

            if (this.Sheets == null) 
            {
                return result;
            }

            result.AddRange(from s in this.Sheets
                            where !string.IsNullOrEmpty(s.PartId)
                            select s.PartId);

            foreach (var s in this.Sheets) 
            {
                if (s.Content != null) 
                {
                    var descendents = s.Content.AllDescendentsOfType<BaseMap>();
                    result.AddRange(from d in descendents
                                    where !string.IsNullOrEmpty(d.PartId)
                                    select d.PartId);
                }
            }

            return result.Distinct();
        }

        internal Sheet GetSheetByInternalId(string id) 
        {
            if (!this.sheetResourceStore.ContainsKey(id)) 
            {
                throw new MetadataException(string.Format("Unknown Sheet.InternalId <{0}>", id));
            }

            var resourceString = this.sheetResourceStore[id];

            using (var sr = new StringReader(resourceString))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (Sheet)XamlReader.Load(xr);
                }
            }
        }

        internal void LoadSheetResources(string resourceString)
        {
            // Loop over each element in the Sheets collection
            // and extract, then add, the xaml which makes up that element.
         
            var document = XDocument.Parse(resourceString);
            foreach (var element in document.Root.Elements())
            {
                if (element.NodeType != XmlNodeType.Element || element.Name == null)
                {
                    continue;
                }

                if (element.Name.LocalName.CompareTo("ExcelDocumentMetadata.Sheets") == 0)
                {
                    int index = 0;

                    foreach (var mapElement in element.Elements())
                    {
                        this.AddSheetResourceStringByIndex(index, mapElement.ToString());
                        index++;
                    }
                }
            }
        }

        private void AddSheetResourceStringByIndex(int index, string resourceString)
        {
            if (this.Sheets.Count > index)
            {
                var sheet = this.Sheets[index];
                if (!this.sheetResourceStore.ContainsKey(sheet.InternalId))
                {
                    this.sheetResourceStore.Add(sheet.InternalId, resourceString);
                }
            }
        }
    }
}

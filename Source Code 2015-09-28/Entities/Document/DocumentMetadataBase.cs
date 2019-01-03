// -----------------------------------------------------------------------
// <copyright file="IDocument.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

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
    /// TODO: Update summary.
    /// </summary>
    public abstract class DocumentMetadataBase
    {
        public DocumentMetadataType DocumentMetadataType { get; internal set; }

        public abstract IEnumerable<string> GetPartIds();

        public bool HasTemplate { get; internal set; }

        /// <summary>
        /// Placeholder for loading an assembly that runs pre process logic.
        /// Must implement interface IDocumentCustomProcess
        /// </summary>
        public string PreProcessAssemblyInfo { get; set; }

        /// <summary>
        /// Placeholder for loading an assembly that runs post process logic.
        /// Must implement interface IDocumentCustomProcess
        /// </summary>
        public string PostProcessAssemblyInfo { get; set; }

        /// <summary>
        /// Byte[] of the physical Template File which can form the starting point of a document
        /// </summary>
        public byte[] TemplateData { get; internal set; }

        /// <summary>
        /// The path of the file that contains template
        /// </summary>
        public string TemplateFileName { get; set; }

        /// <summary>
        /// Deserializes the supplied XML into a <see cref="DocumentMetadataBase"/>
        /// derived entity (i.e. <see cref="ExportMetadata"/> or <see cref="ExcelDocumentMetadata"/>)
        /// </summary>
        internal static DocumentMetadataBase Deserialize(string value)
        {
            DocumentMetadataBase documentMetadata = null;

            using (var sr = new StringReader(value))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    documentMetadata = (DocumentMetadataBase)XamlReader.Load(xr);
                }
            }

            if (documentMetadata is ExportMetadata) 
            {
                documentMetadata.DocumentMetadataType = DocumentMetadataType.ExportMetadata;
            }
            else if (documentMetadata is ExcelDocumentMetadata)
            {
                documentMetadata.DocumentMetadataType = DocumentMetadataType.ExcelDocument;

                var excelDocumentMetadata = (ExcelDocumentMetadata)documentMetadata;
                excelDocumentMetadata.LoadSheetResources(value);
                excelDocumentMetadata.ResourceStore = ResourceStore.Parse(value, null);
            }

            return documentMetadata;
        }

    }
}

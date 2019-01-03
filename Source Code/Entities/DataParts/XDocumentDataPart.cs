using System;
using System.Linq;
using System.Text;
using System.Collections;
using System.Xml.Linq;
using System.Windows.Data;
using System.Xml;

namespace ExcelWriter
{
    public class XDocumentDataPart : IDataPart
    {
        private readonly XDocument document;
        private readonly XmlDataProvider xmlDataProvider;

        public XDocumentDataPart(string partId, XDocument document)
        {
            if (string.IsNullOrEmpty(partId)) 
            {
                throw new ArgumentNullException("partId");
            }
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            this.PartId = partId;
            this.document = document;

            var xmlDocument = new XmlDocument();
            using (var xmlReader = this.document.CreateReader())
            {
                xmlDocument.Load(xmlReader);
            }

            this.xmlDataProvider = new XmlDataProvider
            {
                IsAsynchronous = false,
                Document = xmlDocument
            };
            this.xmlDataProvider.Refresh();
        }

        public ExportParameters ExportParameters { private set; get; }

        public object Data
        {
            get { return this.xmlDataProvider; }
        }

        public int RowCount
        {
            get 
            {
                return 0;
            }
        }

        public string PartId
        {
            get;
            set;
        }
    }

}

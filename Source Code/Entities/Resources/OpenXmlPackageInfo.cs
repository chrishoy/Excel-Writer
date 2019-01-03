namespace ExcelWriter
{
    using System;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;

    public sealed class OpenXmlPackageInfo
    {
        private SpreadsheetDocument package;

        public OpenXmlPackageInfo(string fileName, byte[] data)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("fileName");
            }
            if (data == null)
            {
                throw new ArgumentNullException("data");
            }

            this.FileName = fileName;
            this.Data = data;
        }

        public string FileName { get; private set; }

        public byte[] Data { get; private set; }

        public SpreadsheetDocument Package
        {
            get
            {
                if (this.package == null)
                {
                    this.package = SpreadsheetDocument.Open(new MemoryStream(this.Data), false);
                }
                return this.package;
            }
        }

        public void Flush()
        {
            if (this.package != null)
            {
                this.package.Close();
                this.package = null;
            }
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace ExcelWriter
{
    public sealed class ExcelTemplateFileInfo
    {
        private SpreadsheetDocument spreadsheetDocument;

        public ExcelTemplateFileInfo(string fileName, byte[] data)
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

        public SpreadsheetDocument SpreadsheetDocument 
        {
            get
            {
                if (this.spreadsheetDocument == null)
                {
                    this.spreadsheetDocument = SpreadsheetDocument.Open(new MemoryStream(this.Data), false);
                }
                return this.spreadsheetDocument;
            }
        }

        public void Flush()
        {
            if (this.spreadsheetDocument != null)
            {
                this.spreadsheetDocument.Close();
                this.spreadsheetDocument = null;
            }
        }
    }
}

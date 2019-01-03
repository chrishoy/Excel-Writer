namespace ExcelWriter
{
    using System;

    public class ExportParametersDataPart : IDataPart
    {
        public ExportParametersDataPart(ExportParameters exportParameters)
        {
            if (exportParameters == null)
            {
                throw new ArgumentNullException("exportParameters");
            }

            this.PartId = "Common.Debug";
            this.ExportParameters = exportParameters;
        }

        public ExportParameters ExportParameters { private set; get; }

        public object Data
        {
            get { return this; }
        }

        public int RowCount
        {
            get 
            {
                return this.ExportParameters != null ? this.ExportParameters.NameValueCollection.Count : 0;
            }
        }

        public string PartId
        {
            get;
            set;
        }
    }

}

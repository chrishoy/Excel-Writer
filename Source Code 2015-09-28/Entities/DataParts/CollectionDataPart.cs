namespace ExcelWriter
{
    using System;
    using System.Linq;
    using System.Collections;

    public class CollectionDataPart
    {
        private readonly string partId;
        private readonly IEnumerable data;

        public CollectionDataPart(string partId, IEnumerable data)
        {
            if (string.IsNullOrEmpty(partId))
            {
                throw new ArgumentNullException("partId");
            }

            if (data == null)
            {
                throw new ArgumentNullException("data");
            }

            this.partId = partId;
            this.data = data;
        }

        public object Data
        {
            get { return this.data; }
        }

        public int RowCount
        {
            get { return this.data.Cast<object>().Count(); }
        }

        public string PartId
        {
            get { return this.partId; }
        }
    }

}

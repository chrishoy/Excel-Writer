using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// </summary>
    public sealed class ExportParameters
    {
        public ExportParameters()
        {
            this.NameValueCollection = new List<NameValueParameter>();
        }

        public List<NameValueParameter> NameValueCollection { get; private set; }

        public bool IncludeDebug { get; set; }
        public bool IsReportInPack { get; set; }        
    }
}

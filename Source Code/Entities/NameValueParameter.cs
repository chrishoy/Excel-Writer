using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// </summary>
    public sealed class NameValueParameter
    {
        public NameValueParameter()
        { }

        public decimal GroupId { get; set; }

        public string Name { get; set; }

        public object Value { get; set; }

        public string ReportCode { get; set; }

        public static bool IsNullOrNoValue(NameValueParameter nv)
        {
            if (nv == null || nv.Value == null)
            {
                return true;
            }
            return false;
        }
    }
}

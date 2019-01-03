using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace ExcelWriter
{
    /// <summary>
    /// Dictionary of fonts
    /// </summary>
    internal class NumberFormatsDictionary : Dictionary<string, UInt32Value>
    {
        public KeyValuePair<string, UInt32Value> Find(string numberFormat)
        {
            return this.SingleOrDefault(x => numberFormat.Equals(x.Key));
        }
    }
}

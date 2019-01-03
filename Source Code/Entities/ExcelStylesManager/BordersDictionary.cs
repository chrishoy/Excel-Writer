using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace ExcelWriter
{
    /// <summary>
    /// Dictionary of borders
    /// </summary>
    internal class BordersDictionary : Dictionary<ExcelCellBorderInfo, UInt32Value>
    {
        public KeyValuePair<ExcelCellBorderInfo, UInt32Value> Find(ExcelCellBorderInfo borderInfo)
        {
            return this.SingleOrDefault(x => borderInfo.Equals(x.Key));
        }
    }
}

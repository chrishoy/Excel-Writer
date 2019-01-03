using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace ExcelWriter
{
    /// <summary>
    /// Dictionary of styles
    /// </summary>
    internal class StylesDictionary : Dictionary<ExcelCellStyleInfo, UInt32Value>
    {
        public KeyValuePair<ExcelCellStyleInfo, UInt32Value> Find(ExcelCellStyleInfo cellInfo)
        {
            return this.SingleOrDefault(x => cellInfo.Equals(x.Key));
        }
    }
}

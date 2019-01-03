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
    internal class FontsDictionary : Dictionary<ExcelCellFontInfo, UInt32Value>
    {
        public KeyValuePair<ExcelCellFontInfo, UInt32Value> Find(ExcelCellFontInfo fontInfo)
        {
            return this.SingleOrDefault(x => fontInfo.Equals(x.Key));
        }
    }
}

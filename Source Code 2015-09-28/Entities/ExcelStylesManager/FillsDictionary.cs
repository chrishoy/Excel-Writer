using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace ExcelWriter
{
    /// <summary>
    /// Dictionary of fill colours
    /// </summary>
    internal class FillsDictionary : Dictionary<System.Windows.Media.Color, UInt32Value>
    {
        public KeyValuePair<System.Windows.Media.Color, UInt32Value> Find(System.Windows.Media.Color colour)
        {
            return this.SingleOrDefault(x => colour.Equals(x.Key));
        }
    }
}

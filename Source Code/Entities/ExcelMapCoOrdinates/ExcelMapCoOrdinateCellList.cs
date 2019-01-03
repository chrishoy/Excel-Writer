using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace ExcelWriter
{
    /// <summary>
    /// Represents a list of <see cref="ExcelMapCoOrdinate"/> objects, indexed by strings (i.e. '1,1' is C1R1).
    /// NB! <see cref="CoOrdinate"/> is used because the hash algorithm is fast.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class ExcelMapCoOrdinateCellList : Dictionary<System.Drawing.Point, ExcelMapCoOrdinate>
    {
    }
}

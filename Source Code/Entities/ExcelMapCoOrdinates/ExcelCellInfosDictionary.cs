namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Holds <see cref="ExcelCellInfo"/>s on a Excel Row/Column co-ordinate basis.
    /// </summary>
    internal class ExcelCellInfosDictionary : Dictionary<System.Drawing.Point, ExcelCellInfo>
    {
    }
}

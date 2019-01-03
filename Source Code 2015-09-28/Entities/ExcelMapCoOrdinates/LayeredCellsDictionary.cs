using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Holds layers of <see cref="ExcelMapCoOrdinate"/> derived entities on a Excel Row/Column co-ordinate basis.
    /// </summary>
    internal class LayeredCellsDictionary : Dictionary<System.Drawing.Point, LayeredCellInfo>
    {
        /// <summary>
        /// Update/Insert - Checks if a co-ordinate record exists for the supplied <see cref="System.Drawing.Point"/>.<br/>
        /// If not, creates and adds to this dictionary using the <see cref="System.Drawing.Point"/> as a key,<br/>
        /// then associates adds the <see cref="ExcelMapCoOrdinate"/> with that co-ordinate to an internal list.
        /// </summary>
        /// <param name="idx"></param>
        /// <param name="mapCoOrddinate"></param>
        public void Upsert(System.Drawing.Point coOrdinate, ExcelMapCoOrdinate mapCoOrddinate)
        {
            LayeredCellInfo info;

            if (this.ContainsKey(coOrdinate))
            {
                info = this[coOrdinate];
            }
            else
            {
                info = new LayeredCellInfo();
                this.Add(coOrdinate, info);
            }
            info.LayeredMaps.Add(mapCoOrddinate);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Holds layers of <see cref="ExcelMapCoOrdinate"/> derived entities on a Excel Column index basis.
    /// </summary>
    internal class LayeredColumnsDictionary : Dictionary<uint, LayeredColumnInfo>
    {
        /// <summary>
        /// Update/Insert - Checks if a column index record exists.<br/>
        /// If not, creates and adds to this dictionary using the index as a key,<br/>
        /// then adds the <see cref="ExcelMapCoOrdinate"/> which is associated with that column to an internal list.
        /// </summary>
        /// <param name="idx">The index of the column (Excel column index)</param>
        /// <param name="mapCoOrddinate">The <see cref="ExcelMapCoOrdinate"/> based entity which is participating in this column</param>
        public void Upsert(uint idx, ExcelMapCoOrdinate mapCoOrddinate)
        {
            LayeredColumnInfo info;

            if (this.ContainsKey(idx))
            {
                info = this[idx];
            }
            else
            {
                info = new LayeredColumnInfo();
                this.Add(idx, info);
            }

            info.Maps.Add(mapCoOrddinate);
        }
    }
}

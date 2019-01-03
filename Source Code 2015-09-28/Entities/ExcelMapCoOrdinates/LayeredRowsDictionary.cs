namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Holds layers of <see cref="ExcelMapCoOrdinate"/> derived entities on a Excel Row index basis.
    /// </summary>
    internal class LayeredRowsDictionary : Dictionary<uint, LayeredRowInfo>
    {
        /// <summary>
        /// Update/Insert - Checks if a row index record exists.<br/>
        /// If not, creates and adds to this dictionary using the index as a key,<br/>
        /// then adds the <see cref="ExcelMapCoOrdinate"/> which is associated with that row to an internal list.
        /// </summary>
        /// <param name="idx">The index of the row (Excel row index)</param>
        /// <param name="mapCoOrddinate">The <see cref="ExcelMapCoOrdinate"/> based entity which is participating in this row</param>
        public void Upsert(uint idx, ExcelMapCoOrdinate mapCoOrddinate)
        {
            LayeredRowInfo info;

            if (this.ContainsKey(idx))
            {
                info = this[idx];
            }
            else
            {
                info = new LayeredRowInfo();
                this.Add(idx, info);
            }

            info.Maps.Add(mapCoOrddinate);
        }
    }
}

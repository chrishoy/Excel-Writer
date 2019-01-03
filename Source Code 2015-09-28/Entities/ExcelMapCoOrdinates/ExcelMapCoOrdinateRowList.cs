using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Reprsents a list of <see cref="ExcelMapCoOrdinateRow"/>s.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class ExcelMapCoOrdinateRowList : Dictionary<uint, ExcelMapCoOrdinateRow>
    {
        /// <summary>
        /// Returns the row in the list that has been marked as having a specified worksheet row index.
        /// </summary>
        /// <param name="worksheetColumnIndex"></param>
        /// <returns></returns>
        public ExcelMapCoOrdinateRow FindRow(uint worksheetRowIndex)
        {
            var row = this.FirstOrDefault(c => c.Value.WorksheetRowIndex == worksheetRowIndex);
            return row.Value;
        }
    }
}

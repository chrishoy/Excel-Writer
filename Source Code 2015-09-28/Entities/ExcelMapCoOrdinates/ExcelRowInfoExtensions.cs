namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Extension methods for <see cref="ExcelRowInfo"/>.
    /// </summary>
    internal static class ExcelRowInfoExtensions
    {
        /// <summary>
        /// Compares and updates this information based on supplied Row
        /// </summary>
        /// <param name="rowInfo">The <see cref="ExcelRowInfo"/> that will be updated.</param>
        /// <param name="row">The <see cref="ExcelMapCoOrdinateRow"/> which will be used to update the <see cref="ExcelRowInfo"/>.</param>
        internal static void Update(this ExcelRowInfo rowInfo, ExcelMapCoOrdinateRow row)
        {
            if (row == null)
            {
                return; // No information to update
            }

            rowInfo.UpdateHasHiddenRow(row.IsHidden);
            rowInfo.UpdateMaxRowHeight(row.Height);
        }

        /// <summary>
        /// Compares and updates the maximum Row Height encountered.
        /// </summary>
        /// <param name="rowInfo">The <see cref="ExcelRowInfo"/> that will be updated.</param>
        /// <param name="rowHeight">Height of attempted row set.</param>
        internal static void UpdateMaxRowHeight(this ExcelRowInfo rowInfo, double? rowHeight)
        {
            // Determine the maximum encountered Row Height
            if (rowHeight.HasValue && rowInfo.MaxRowHeight.GetValueOrDefault() < rowHeight.Value)
            {
                rowInfo.MaxRowHeight = rowHeight;
                rowInfo.HasRowInfo = true;
            }
        }

        /// <summary>
        /// Compares and updates the HasHiddenRow property.
        /// </summary>
        /// <param name="rowInfo">The <see cref="ExcelRowInfo"/> that will be updated.</param>
        /// <param name="rowHidden">True if row is to be hidden, false otherwise.</param>
        internal static void UpdateHasHiddenRow(this ExcelRowInfo rowInfo, bool rowHidden)
        {
            if (!rowInfo.HasHiddenRow)
            {
                rowInfo.HasHiddenRow = rowHidden;
                rowInfo.HasRowInfo = true;
            }
        }
    }
}

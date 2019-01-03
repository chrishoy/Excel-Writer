// -----------------------------------------------------------------------
// <copyright file="ExcelCellBorderInfoExtensions.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    internal static class ExcelCellBorderInfoExtensions
    {
        /// <summary>
        /// Extension methods that can be used with <see cref="ExcelCellBorderInfo"/>.
        /// </summary>
        /// <param name="rowHeight"></param>
        public static void UpdateBorder(this ExcelCellBorderInfo borderToUpdate, ExcelCellBorderInfo borderInfoToApply)
        {
            // Anything to apply?
            if (borderInfoToApply == null || borderInfoToApply.HasBorder == false) return;

            // Apply updates to the thickness (ie. any non-zero values)
            if (borderInfoToApply.HasLeftBorder)
            {
                borderToUpdate.WidthLeft = borderInfoToApply.WidthLeft;
                borderToUpdate.ColourLeft = borderInfoToApply.ColourLeft;
            }

            if (borderInfoToApply.HasTopBorder)
            {
                borderToUpdate.WidthTop = borderInfoToApply.WidthTop;
                borderToUpdate.ColourTop = borderInfoToApply.ColourTop;
            }

            if (borderInfoToApply.HasRightBorder)
            {
                borderToUpdate.WidthRight = borderInfoToApply.WidthRight;
                borderToUpdate.ColourRight = borderInfoToApply.ColourRight;
            }

            if (borderInfoToApply.HasBottomBorder)
            {
                borderToUpdate.WidthBottom = borderInfoToApply.WidthBottom;
                borderToUpdate.ColourBottom = borderInfoToApply.ColourBottom;
            }
        }

        /// <summary>
        /// Applies border on top of this border.<br/>
        /// Where the frame thickness is 0, the underlying border is not updated,<br/>
        /// otherwise it is over-written with a new border, using the applied colour.
        /// </summary>
        /// <param name="rowHeight"></param>
        public static void UpdateBorder(this ExcelCellBorderInfo borderToUpdate, StyleBase styleToApply)
        {
            // Anything to apply?
            if (styleToApply == null || styleToApply.HasAnyBorder() == false) return;

            // Apply updates to the thickness (ie. any non-zero values)
            if (styleToApply.HasLeftBorder())
            {
                borderToUpdate.WidthLeft = styleToApply.BorderThickness.Value.Left;
                borderToUpdate.ColourLeft = styleToApply.BorderColour;
            }

            if (styleToApply.HasTopBorder())
            {
                borderToUpdate.WidthTop = styleToApply.BorderThickness.Value.Top;
                borderToUpdate.ColourTop = styleToApply.BorderColour;
            }

            if (styleToApply.HasRightBorder())
            {
                borderToUpdate.WidthRight = styleToApply.BorderThickness.Value.Right;
                borderToUpdate.ColourRight = styleToApply.BorderColour;
            }

            if (styleToApply.HasBottomBorder())
            {
                borderToUpdate.WidthBottom = styleToApply.BorderThickness.Value.Bottom;
                borderToUpdate.ColourBottom = styleToApply.BorderColour;
            }
        }
    }
}

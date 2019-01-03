// -----------------------------------------------------------------------
// <copyright file="StyleBaseExtensions.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows.Media;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    internal static class StyleBaseExtensions
    {
        /// <summary>
        /// Checks the <see cref="StyleBase"/> BorderColour and BorderThickness values to determine if there is any
        /// border information to be applied to the <see cref="StyleBase"/>.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        public static bool HasAnyBorder(this StyleBase mapStyle)
        {
            //if (mapStyle == null) return false;

            if (mapStyle.BorderThickness == null || !mapStyle.BorderThickness.HasValue) return false;
            if (mapStyle.BorderColour == null || !mapStyle.BorderColour.HasValue || mapStyle.BorderColour.Value == Colors.Transparent) return false;

            if (mapStyle.BorderThickness.Value.Left == 0 &&
                mapStyle.BorderThickness.Value.Top == 0 &&
                mapStyle.BorderThickness.Value.Right == 0 &&
                mapStyle.BorderThickness.Value.Bottom == 0) return false;

            return true;
        }

        /// <summary>
        /// Checks the <see cref="StyleBase"/> BorderColour and BorderThickness values to determine if there is any
        /// border information to be applied to the <see cref="StyleBase"/>.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        public static bool HasLeftBorder(this StyleBase mapStyle)
        {
            //if (mapStyle == null) return false;

            if (mapStyle.BorderThickness == null || !mapStyle.BorderThickness.HasValue) return false;
            if (mapStyle.BorderColour == null || !mapStyle.BorderColour.HasValue || mapStyle.BorderColour.Value == Colors.Transparent) return false;

            return mapStyle.BorderThickness.Value.Left > 0;
        }

        /// <summary>
        /// Checks the <see cref="StyleBase"/> BorderColour and BorderThickness values to determine if there is any
        /// border information to be applied to the <see cref="StyleBase"/>.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        public static bool HasTopBorder(this StyleBase mapStyle)
        {
            //if (mapStyle == null) return false;

            if (mapStyle.BorderThickness == null || !mapStyle.BorderThickness.HasValue) return false;
            if (mapStyle.BorderColour == null || !mapStyle.BorderColour.HasValue || mapStyle.BorderColour.Value == Colors.Transparent) return false;

            return mapStyle.BorderThickness.Value.Top > 0;
        }

        /// <summary>
        /// Checks the <see cref="StyleBase"/> BorderColour and BorderThickness values to determine if there is any
        /// border information to be applied to the <see cref="StyleBase"/>.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        public static bool HasRightBorder(this StyleBase mapStyle)
        {
            //if (mapStyle == null) return false;

            if (mapStyle.BorderThickness == null || !mapStyle.BorderThickness.HasValue) return false;
            if (mapStyle.BorderColour == null || !mapStyle.BorderColour.HasValue || mapStyle.BorderColour.Value == Colors.Transparent) return false;

            return mapStyle.BorderThickness.Value.Right > 0;
        }

        /// <summary>
        /// Checks the <see cref="StyleBase"/> BorderColour and BorderThickness values to determine if there is any
        /// border information to be applied to the <see cref="StyleBase"/>.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        public static bool HasBottomBorder(this StyleBase mapStyle)
        {
            //if (mapStyle == null) return false;

            if (mapStyle.BorderThickness == null || !mapStyle.BorderThickness.HasValue) return false;
            if (mapStyle.BorderColour == null || !mapStyle.BorderColour.HasValue || mapStyle.BorderColour.Value == Colors.Transparent) return false;

            return mapStyle.BorderThickness.Value.Bottom > 0;
        }

    }
}

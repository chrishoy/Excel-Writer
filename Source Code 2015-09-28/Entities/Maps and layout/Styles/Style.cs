using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Windows;

namespace ExcelWriter
{
    /// <summary>
    /// Style attributes which can be specifically used with <see cref="Map"/>s.
    /// </summary>
    public class Style : StyleBase
    {
        /// <summary>
        /// Create and return a copy of this instance.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return new Style
            {
                BackgroundColour = this.BackgroundColour,
                BasedOnKey = this.BasedOnKey,
                BorderColour = this.BorderColour,
                BorderThickness = this.BorderThickness,
                Key = this.Key
            };
        }

        /// <summary>
        /// Creates a new style which is based on an existing style, merging over current values.
        /// </summary>
        /// <param name="basedOnStyle">Style on which the new style is based</param>
        /// <param name="styleToMerge">Style which is to be berged over the base style</param>
        /// <returns></returns>
        public static Style CreateMergedStyle(StyleBase basedOnStyle, Style styleToMerge)
        {
            // First clone
            var newStyle = (Style)styleToMerge.Clone();

            // Then override
            newStyle.BackgroundColour = styleToMerge.BackgroundColour.HasValue ? styleToMerge.BackgroundColour : basedOnStyle.BackgroundColour;
            newStyle.BorderColour = styleToMerge.BorderColour.HasValue ? styleToMerge.BorderColour : basedOnStyle.BorderColour;
            newStyle.BorderThickness = styleToMerge.BorderThickness.HasValue ? styleToMerge.BorderThickness : basedOnStyle.BorderThickness;

            return newStyle;
        }

    }
}

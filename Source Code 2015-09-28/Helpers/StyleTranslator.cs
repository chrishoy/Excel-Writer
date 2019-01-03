namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Media;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;

    using OpenXml.Excel;

    internal sealed partial class StyleTranslator
    {

        public static BorderStyleValues Translate(double thickness)
        {
            if (thickness == 0)
            {
                return BorderStyleValues.None;
            }
            else if (thickness == 1)
            {
                return BorderStyleValues.Thin;
            }
            else if (thickness == 2)
            {
                return BorderStyleValues.Medium;
            }

            return BorderStyleValues.Thick;
        }

        public static HorizontalAlignmentValues Translate(TextAlignment textAlignment)
        {
            switch (textAlignment)
            {
                case TextAlignment.Center:
                    return HorizontalAlignmentValues.Center;
                case TextAlignment.Justify:
                    return HorizontalAlignmentValues.Justify;
                case TextAlignment.Right:
                    return HorizontalAlignmentValues.Right;
                default:
                    return HorizontalAlignmentValues.Left;
            }
        }

        public static VerticalAlignmentValues Translate(VerticalAlignment textAlignment)
        {
            switch (textAlignment)
            {
                case VerticalAlignment.Center:
                    return VerticalAlignmentValues.Center;
                case VerticalAlignment.Stretch:
                    return VerticalAlignmentValues.Justify;
                case VerticalAlignment.Top:
                    return VerticalAlignmentValues.Top;
                default:
                    return VerticalAlignmentValues.Bottom;
            }
        }

        public static string Translate(System.Windows.Media.Color color)
        {
            return System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb(
                                color.A,
                                color.R,
                                color.G,
                                color.B)).Replace("#", string.Empty);
        }

    }
}

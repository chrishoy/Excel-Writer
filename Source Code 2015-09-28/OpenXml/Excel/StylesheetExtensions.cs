using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelWriter.OpenXml.Excel
{
    public static class StylesheetExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Adds the cell format to the stylesheet.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        /// <param name="cellFormat">The cell format.</param>
        /// <returns>The cell format index.</returns>
        public static uint AddCellFormat(this Stylesheet stylesheet, CellFormat cellFormat)
        {
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
            }

            stylesheet.CellFormats.Append(cellFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            uint cellFormatId = (uint)stylesheet.CellFormats.ToList().IndexOf(cellFormat);

            return cellFormatId;
        }

        public static uint AddDifferentialFormat(this Stylesheet stylesheet, DifferentialFormat df)
        {
            if (stylesheet.DifferentialFormats == null)
            {
                stylesheet.DifferentialFormats = new DifferentialFormats();
            }
        
            stylesheet.DifferentialFormats.Append(df);
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Count();
            uint dfId = (uint)stylesheet.DifferentialFormats.ToList().IndexOf(df);

            return dfId;
        }

        #region Border

        public static bool TryMatchBorder(this Stylesheet stylesheet, Border candidate, out uint id)
        {
            id = 0;
            var match = (from f in stylesheet.Borders
                         where f.InnerXml.CompareTo(candidate.InnerXml) == 0
                         select f).FirstOrDefault();

            if (match != null)
            {
                id = (uint)stylesheet.Borders.ToList().IndexOf(match);
                return true;
            }

            return false;
        }

        public static uint AddBorder(this Stylesheet stylesheet, Border border)
        {
            Border clone = (Border)border.Clone();
            stylesheet.Borders.Append(clone);
            return (uint)stylesheet.Borders.ToList().IndexOf(clone);
        }

        #endregion

        #region Fills

        public static bool TryMatchFill(this Stylesheet stylesheet, Fill candidate, out uint id)
        {
            id = 0;
            var match = (from f in stylesheet.Fills
                         where f.InnerXml.CompareTo(candidate.InnerXml) == 0
                         select f).FirstOrDefault();

            if (match != null)
            {
                id = (uint)stylesheet.Fills.ToList().IndexOf(match);
                return true;
            }

            return false;
        }

        public static uint AddFill(this Stylesheet stylesheet, Fill fill)
        {
            Fill clone = (Fill)fill.Clone();
            stylesheet.Fills.Append(clone);
            return (uint)stylesheet.Fills.ToList().IndexOf(clone);
        }

        /// <summary>
        /// Adds the fill to the stylesheet.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <returns>The fill index.</returns>
        public static uint AddFill(this Stylesheet stylesheet, System.Windows.Media.Color foregroundColor)
        {
            return AddFill(stylesheet, foregroundColor, PatternValues.Solid);
        }

        /// <summary>
        /// Adds the fill to the stylesheet.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <param name="patternValue">The pattern value.</param>
        /// <returns>The fill index.</returns>
        public static uint AddFill(this Stylesheet stylesheet, System.Windows.Media.Color foregroundColor, PatternValues patternValue)
        {
            Fill fill = new Fill();
            fill.PatternFill = new PatternFill()
            {
                ForegroundColor = new ForegroundColor()
                {
                    Rgb = new HexBinaryValue(TranslateColor(foregroundColor))
                },
                PatternType = new EnumValue<PatternValues>(patternValue),
            };
            stylesheet.Fills.Append(fill);
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
            uint fillId = (uint)stylesheet.Fills.ToList().IndexOf(fill);

            return fillId;
        }

        public static uint AddIndexedColor(this Stylesheet stylesheet, System.Windows.Media.Color color)
        {
            uint fillId = stylesheet.AddFill(color);
            string hexColor = "00" + TranslateColor(color);

            if (stylesheet.Colors == null)
            {
                stylesheet.Colors = new Colors();
            }

            if (stylesheet.Colors.IndexedColors == null)
            {
                stylesheet.Colors.IndexedColors = new IndexedColors();
            }

            var indexedColor = stylesheet.Colors.IndexedColors.FirstOrDefault(x => string.Equals(((RgbColor)x).Rgb.Value, hexColor));
            if (indexedColor == null)
            {
                stylesheet.Colors.IndexedColors.Append(
                    new RgbColor()
                    {
                        Rgb = hexColor
                    });
                //var elementAt = stylesheet.Colors.IndexedColors.ElementAt(stylesheet.Colors.IndexedColors.Count() - 10);
                //elementAt.Remove();
            }

            return fillId;
        }

        #endregion

        #region Fonts

        public static bool TryMatchFont(this Stylesheet stylesheet, Font candidate, out uint id)
        {
            id = 0;
            var match = (from f in stylesheet.Fonts
                         where f.InnerXml.CompareTo(candidate.InnerXml) == 0
                         select f).FirstOrDefault();

            if (match != null)
            {
                id = (uint)stylesheet.Fonts.ToList().IndexOf(match);
                return true;
            }
            return false;
        }

        public static uint AddFont(this Stylesheet stylesheet, Font font)
        {
            Font clone = (Font)font.Clone();
            stylesheet.Fonts.Append(clone);
            return (uint)stylesheet.Fonts.ToList().IndexOf(clone);
        }

        /// <summary>
        /// Adds the font to the stylesheet.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        /// <param name="fontName">Name of the font.</param>
        /// <param name="fontSize">Size of the font.</param>
        /// <param name="isBold">if set to <c>true</c> bold.</param>
        /// <param name="isItalic">if set to <c>true</c> italic.</param>
        /// <param name="isUnderline">if set to <c>true</c> underline.</param>
        /// <param name="foregroundColor">Color of the foreground.</param>
        /// <returns>The font index.</returns>
        public static uint AddFont(
            this Stylesheet stylesheet,
            string fontName,
            double fontSize,
            bool isBold,
            bool isItalic,
            bool isUnderline,
            System.Windows.Media.Color foregroundColor)
        {
            Font font = new Font()
            {
                Color = new Color()
                {
                    Rgb = new HexBinaryValue(TranslateColor(foregroundColor))
                },
                FontName = new FontName()
                {
                    Val = new StringValue(fontName)
                },
                FontSize = new FontSize()
                {
                    Val = new DoubleValue(fontSize)
                }
            };

            if (isBold)
            {
                font.Bold = new Bold();
            }

            if (isItalic)
            {
                font.Italic = new Italic();
            }

            if (isUnderline)
            {
                font.Underline = new Underline();
            }

            stylesheet.Fonts.Append(font);
            uint fontId = (uint)stylesheet.Fonts.ToList().IndexOf(font);

            return fontId;
        }

        #endregion

        #region Numbering formats

        public static bool TryMatchNumberingFormat(this Stylesheet stylesheet, NumberingFormat candidate, out uint id)
        {
            id = 0;

            if (stylesheet.NumberingFormats == null || candidate == null || !candidate.FormatCode.HasValue) 
            {
                return false;
            }

            var match = (from f in stylesheet.NumberingFormats.Descendants<NumberingFormat>()
                         where f.FormatCode != null && f.FormatCode.HasValue && f.FormatCode.Value.CompareTo(candidate.FormatCode.Value) == 0
                         select f).FirstOrDefault();

            if (match != null && match.NumberFormatId.HasValue)
            {
                id = match.NumberFormatId.Value;
                return true;
            }
            return false;
        }

        public static uint AddNumberingFormat(this Stylesheet stylesheet, NumberingFormat fill)
        {
            NumberingFormat clone = (NumberingFormat)fill.Clone();

            if (stylesheet.NumberingFormats == null) 
            {
                stylesheet.NumberingFormats = new NumberingFormats();
            }
            stylesheet.NumberingFormats.Append(clone);
            var id = 164 + (uint)stylesheet.NumberingFormats.ToList().IndexOf(clone);
            clone.NumberFormatId = id;
            return id;
        }

        /// <summary>
        /// Adds the numbering format.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        /// <param name="formatCode">The format code.</param>
        /// <returns>the numbering format index.</returns>
        public static uint AddNumberingFormat(this Stylesheet stylesheet, string formatCode)
        {
            NumberingFormat numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = formatCode;

            if (stylesheet.NumberingFormats == null)
            {
                stylesheet.NumberingFormats = new NumberingFormats();
            }


            // NB. Very important, custom numbering formats must start from 164
            stylesheet.NumberingFormats.Append(numberingFormat);
            stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            uint numberingFormatId = 164 + (uint)stylesheet.NumberingFormats.ToList().IndexOf(numberingFormat);
            numberingFormat.NumberFormatId = numberingFormatId;

            return numberingFormatId;
        }

        public static void InitializeIndexedColors(this Stylesheet stylesheet)
        {            
            if (stylesheet.Colors == null)
            {
                stylesheet.Colors = new Colors();
            }
            if (stylesheet.Colors.IndexedColors == null)
            {
                stylesheet.Colors.IndexedColors = new IndexedColors();
            }

            IndexedColors indexedColors = stylesheet.Colors.IndexedColors;
            indexedColors.RemoveAllChildren<RgbColor>();

            // GAM Excel 2003 Colours
            RgbColor rgbColor1 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor2 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor3 = new RgbColor() { Rgb = "00FF0000" };
            RgbColor rgbColor4 = new RgbColor() { Rgb = "0000FF00" };
            RgbColor rgbColor5 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor6 = new RgbColor() { Rgb = "00FFFF00" };
            RgbColor rgbColor7 = new RgbColor() { Rgb = "00FF00FF" };
            RgbColor rgbColor8 = new RgbColor() { Rgb = "0000FFFF" };
            RgbColor rgbColor9 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor10 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor11 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor12 = new RgbColor() { Rgb = "009EBEE8" };
            RgbColor rgbColor13 = new RgbColor() { Rgb = "00BBB6B4" };
            RgbColor rgbColor14 = new RgbColor() { Rgb = "00D8BC8F" };
            RgbColor rgbColor15 = new RgbColor() { Rgb = "00AAB490" };
            RgbColor rgbColor16 = new RgbColor() { Rgb = "00D797C3" };
            RgbColor rgbColor17 = new RgbColor() { Rgb = "00E3A197" };
            RgbColor rgbColor18 = new RgbColor() { Rgb = "00D7CD91" };
            RgbColor rgbColor19 = new RgbColor() { Rgb = "00AF9BBF" };
            RgbColor rgbColor20 = new RgbColor() { Rgb = "0066CAB9" };
            RgbColor rgbColor21 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor22 = new RgbColor() { Rgb = "0066A3A3" };
            RgbColor rgbColor23 = new RgbColor() { Rgb = "00B2D1D1" };
            RgbColor rgbColor24 = new RgbColor() { Rgb = "00F2F7F7" };
            RgbColor rgbColor25 = new RgbColor() { Rgb = "009EBEE8" };
            RgbColor rgbColor26 = new RgbColor() { Rgb = "00BBB6B4" };
            RgbColor rgbColor27 = new RgbColor() { Rgb = "00D8BC8F" };
            RgbColor rgbColor28 = new RgbColor() { Rgb = "00AAB490" };
            RgbColor rgbColor29 = new RgbColor() { Rgb = "00D797C3" };
            RgbColor rgbColor30 = new RgbColor() { Rgb = "00E3A197" };
            RgbColor rgbColor31 = new RgbColor() { Rgb = "00D7CD91" };
            RgbColor rgbColor32 = new RgbColor() { Rgb = "00AF9BBF" };
            RgbColor rgbColor33 = new RgbColor() { Rgb = "0066CAB9" };
            RgbColor rgbColor34 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor35 = new RgbColor() { Rgb = "0066A3A3" };
            RgbColor rgbColor36 = new RgbColor() { Rgb = "00B2D1D1" };
            RgbColor rgbColor37 = new RgbColor() { Rgb = "00F2F7F7" };
            RgbColor rgbColor38 = new RgbColor() { Rgb = "00800000" };
            RgbColor rgbColor39 = new RgbColor() { Rgb = "00008080" };
            RgbColor rgbColor40 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor41 = new RgbColor() { Rgb = "0000CCFF" };
            RgbColor rgbColor42 = new RgbColor() { Rgb = "00CCFFFF" };
            RgbColor rgbColor43 = new RgbColor() { Rgb = "00CCFFCC" };
            RgbColor rgbColor44 = new RgbColor() { Rgb = "00FFFF99" };
            RgbColor rgbColor45 = new RgbColor() { Rgb = "0099CCFF" };
            RgbColor rgbColor46 = new RgbColor() { Rgb = "00FF99CC" };
            RgbColor rgbColor47 = new RgbColor() { Rgb = "00CC99FF" };
            RgbColor rgbColor48 = new RgbColor() { Rgb = "00FFCC99" };
            RgbColor rgbColor49 = new RgbColor() { Rgb = "003366FF" };
            RgbColor rgbColor50 = new RgbColor() { Rgb = "0033CCCC" };
            RgbColor rgbColor51 = new RgbColor() { Rgb = "0099CC00" };
            RgbColor rgbColor52 = new RgbColor() { Rgb = "00FFCC00" };
            RgbColor rgbColor53 = new RgbColor() { Rgb = "00FF9900" };
            RgbColor rgbColor54 = new RgbColor() { Rgb = "00FF6600" };
            RgbColor rgbColor55 = new RgbColor() { Rgb = "00666699" };
            RgbColor rgbColor56 = new RgbColor() { Rgb = "00969696" };
            RgbColor rgbColor57 = new RgbColor() { Rgb = "00003366" };
            RgbColor rgbColor58 = new RgbColor() { Rgb = "00339966" };
            RgbColor rgbColor59 = new RgbColor() { Rgb = "00003300" };
            RgbColor rgbColor60 = new RgbColor() { Rgb = "00333300" };
            RgbColor rgbColor61 = new RgbColor() { Rgb = "00993300" };
            RgbColor rgbColor62 = new RgbColor() { Rgb = "00993366" };
            RgbColor rgbColor63 = new RgbColor() { Rgb = "00333399" };
            RgbColor rgbColor64 = new RgbColor() { Rgb = "00333333" };
            indexedColors.Append(rgbColor1);
            indexedColors.Append(rgbColor2);
            indexedColors.Append(rgbColor3);
            indexedColors.Append(rgbColor4);
            indexedColors.Append(rgbColor5);
            indexedColors.Append(rgbColor6);
            indexedColors.Append(rgbColor7);
            indexedColors.Append(rgbColor8);
            indexedColors.Append(rgbColor9);
            indexedColors.Append(rgbColor10);
            indexedColors.Append(rgbColor11);
            indexedColors.Append(rgbColor12);
            indexedColors.Append(rgbColor13);
            indexedColors.Append(rgbColor14);
            indexedColors.Append(rgbColor15);
            indexedColors.Append(rgbColor16);
            indexedColors.Append(rgbColor17);
            indexedColors.Append(rgbColor18);
            indexedColors.Append(rgbColor19);
            indexedColors.Append(rgbColor20);
            indexedColors.Append(rgbColor21);
            indexedColors.Append(rgbColor22);
            indexedColors.Append(rgbColor23);
            indexedColors.Append(rgbColor24);
            indexedColors.Append(rgbColor25);
            indexedColors.Append(rgbColor26);
            indexedColors.Append(rgbColor27);
            indexedColors.Append(rgbColor28);
            indexedColors.Append(rgbColor29);
            indexedColors.Append(rgbColor30);
            indexedColors.Append(rgbColor31);
            indexedColors.Append(rgbColor32);
            indexedColors.Append(rgbColor33);
            indexedColors.Append(rgbColor34);
            indexedColors.Append(rgbColor35);
            indexedColors.Append(rgbColor36);
            indexedColors.Append(rgbColor37);
            indexedColors.Append(rgbColor38);
            indexedColors.Append(rgbColor39);
            indexedColors.Append(rgbColor40);
            indexedColors.Append(rgbColor41);
            indexedColors.Append(rgbColor42);
            indexedColors.Append(rgbColor43);
            indexedColors.Append(rgbColor44);
            indexedColors.Append(rgbColor45);
            indexedColors.Append(rgbColor46);
            indexedColors.Append(rgbColor47);
            indexedColors.Append(rgbColor48);
            indexedColors.Append(rgbColor49);
            indexedColors.Append(rgbColor50);
            indexedColors.Append(rgbColor51);
            indexedColors.Append(rgbColor52);
            indexedColors.Append(rgbColor53);
            indexedColors.Append(rgbColor54);
            indexedColors.Append(rgbColor55);
            indexedColors.Append(rgbColor56);
            indexedColors.Append(rgbColor57);
            indexedColors.Append(rgbColor58);
            indexedColors.Append(rgbColor59);
            indexedColors.Append(rgbColor60);
            indexedColors.Append(rgbColor61);
            indexedColors.Append(rgbColor62);
            indexedColors.Append(rgbColor63);
            indexedColors.Append(rgbColor64);
        }

        #endregion

        #endregion

        #region Private Static Methods

        private static string TranslateColor(System.Windows.Media.Color color)
        {
            return System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb(
                                color.A,
                                color.R,
                                color.G,
                                color.B)).Replace("#", string.Empty);
        }

        #endregion
    }
}

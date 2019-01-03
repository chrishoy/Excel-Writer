namespace ExcelWriter
{
    using System;
    using System.Drawing;

    /// <summary>
    /// Converts column width and how heights in Excel to and from pixels.<br/>
    /// This is extraordinarily (and probably unnecessarily) complicated process which has had to be been empirically derived<br/>
    /// as, at the time of writing, Microsoft don't appear to be able to supply any concrete documentation about this.
    /// </summary>
    public class ExcelDimensionConverter
    {
        /// <summary>
        /// The standard conversion seems to be based on 96 pixels per inch, 72 points per inch, therefore 96 pixels per 72 points or 1 pixel per 0.75 points.
        /// </summary>
        private const double POINTS_PER_PIXEL = OpenXml.Excel.Constants.PointsPerInch / 96;

        /// <summary>
        /// Padding is crucial in column width calculations. It is 1 + 2 + 2 (border + start + end padding) 
        /// </summary>
        private const int PADDING = 5;
        
        /// <summary>
        /// This is the maximum width of the 'Normal' font characters 0 to 9
        /// </summary>
        private int maxCharWidthInPixels;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelDimensionConverter" /> class.
        /// </summary>
        /// <param name="fontFamily">The font family</param>
        /// <param name="size">The size of the font</param>
        public ExcelDimensionConverter(string fontFamily, float size)
        {
            int maxCharWidthInPixels = (int)this.CalculateMaxDigitWidth(fontFamily, size);
            this.maxCharWidthInPixels = maxCharWidthInPixels;
        }

        #region Public Methods

        /// <summary>
        /// Converts a width, set in OpenXML to the widths that we will actually get in the final Excel document.<br/>
        /// The result will be fractionally smaller. This is required purely because Microsoft likes to make life very hard for us...!
        /// </summary>
        /// <param name="width">The Width we would otherwise set in OpenXML</param>
        /// <returns>The width that an Excel column based on the OpenXML width would actually be.</returns>
        public double OpenXmlWidthToWidth(double width)
        {
            return this.PixelsToWidth(this.OpenXmlWidthToPixels(width));
        }

        /// <summary>
        /// Converts a width in an Excel document to the width that we would have to specify in OpenXML to get that width.<br/>
        /// The result will be fractionally larger. This is required purely because Microsoft likes to make life very hard for us...!
        /// </summary>
        /// <param name="width">The width of the column in Excel that we actually want.</param>
        /// <returns>The width that we would have to specify in OpenXML to get that width.</returns>
        public double WidthToOpenXmlWidth(double width)
        {
            return this.PixelsToOpenXmlWidth(this.WidthToPixels(width));
        }

        /// <summary>
        /// Converts a height that we would set in OpenXML to the a height that we would actually get in an Excel document.<br/>
        /// Why Width is not the same I have absolutely no idea...!
        /// </summary>
        /// <param name="height">The Height we would otherwise set in OpenXML</param>
        /// <returns>The height that an Excel row based on the OpenXML width would actually be.</returns>
        public double OpenXmlHeightToHeight(double height)
        {
            return this.PixelsToHeight(this.OpenXmlHeightToPixels(height));
        }

        /// <summary>
        /// Converts a height in an Excel document to the height that we would have to specify in OpenXML to get that height.<br/>
        /// Why Width is not the same I have absolutely no idea...!
        /// </summary>
        /// <param name="height">The height of the row in Excel that we actually want.</param>
        /// <returns>The height that we would have to specify in OpenXML to get that height.</returns>
        public double HeightToOpenXmlHeight(double height)
        {
            return this.PixelsToOpenXmlHeight(this.HeightToPixels(height));
        }

        /// <summary>
        /// Converts a height (as specified in an Excel row) to a number of pixels.
        /// </summary>
        /// <param name="height">The excel column height</param>
        /// <returns>The number of pixels</returns>
        public int HeightToPixels(double height)
        {
            return (int)(height / POINTS_PER_PIXEL);
        }

        /// <summary>
        /// Converts a supplied number of pixels to a row height in Excel.
        /// </summary>
        /// <param name="pixels">The number of pixels</param>
        /// <returns>A row height in Excel</returns>
        public double PixelsToHeight(int pixels)
        {
            return pixels * POINTS_PER_PIXEL;
        }

        /// <summary>
        /// Converts English Metric Units to Pixels
        /// </summary>
        /// <param name="emus">Number of 'English Metric Units'</param>
        /// <returns>Number of pixels</returns>
        public int EmusToPixels(long emus)
        {
            long points = emus / (long)OpenXml.Excel.Constants.EmusPerPoint;
            return (int)(points / POINTS_PER_PIXEL);
        }

        /// <summary>
        /// Converts Pixels to English Metric Units
        /// </summary>
        /// <param name="pixels">Number of pixels</param>
        /// <returns>Number of 'English Metric Units'</returns>
        public long PixelsToEmus(long pixels)
        {
            double points = pixels * POINTS_PER_PIXEL;
            return (long)(points * OpenXml.Excel.Constants.EmusPerPoint);
        }

        /// <summary>
        /// Converts a supplied OpenXML height to an integer number of pixels (the lowest unit in Excel)
        /// </summary>
        /// <param name="width">The OpenXML height</param>
        /// <returns>The nearest equivalent number of pixels</returns>
        public int OpenXmlHeightToPixels(double height)
        {
            // No trimming/padding occurs with height
            return HeightToPixels(height);
        }

        /// <summary>
        /// Converts a supplied number of pixels to a height in OpenXML
        /// </summary>
        /// <param name="pixels">The number of pixels</param>
        /// <returns>A height in OpenXML</returns>
        public double PixelsToOpenXmlHeight(int pixels)
        {
            // No trimming/padding occurs with height
            return PixelsToHeight(pixels);
        }

        /// <summary>
        /// Converts a supplied OpenXML width to an integer number of pixels (the lowest unit in Excel)
        /// </summary>
        /// <param name="width">The OpenXML width</param>
        /// <returns>The nearest equivalent number of pixels</returns>
        public int OpenXmlWidthToPixels(double width)
        {
            return (int)Math.Truncate((((256d * width) + Math.Truncate(128d / this.maxCharWidthInPixels)) / 256d) * this.maxCharWidthInPixels); 
        }

        /// <summary>
        /// Converts a width (as specified in an Excel column) to a number of pixels.
        /// </summary>
        /// <param name="width">The excel column width</param>
        /// <returns>The number of pixels</returns>
        public int WidthToPixels(double width)
        {
            double adjustedPadding = width >= 1 ? PADDING : PADDING * width;
            double widthExPadding = Math.Truncate((((256d * width) + Math.Truncate(128d / this.maxCharWidthInPixels)) / 256d) * this.maxCharWidthInPixels);
            return (int)Math.Round(widthExPadding + adjustedPadding, 0); 
        }

        /// <summary>
        /// Converts a supplied number of pixels to a width in Excel.
        /// </summary>
        /// <param name="pixels">The number of pixels</param>
        /// <returns>A width in Excel</returns>
        public double PixelsToWidth(int pixels)
        {
            // Determine the actual padding. This is 5 (the padding) until the number 
            // of pixels becomes lower than the 'maximum character width' + 5
            double actualPadding = ((pixels - PADDING) < this.maxCharWidthInPixels)
                ? PADDING * ((double)pixels / (PADDING + this.maxCharWidthInPixels))
                : PADDING;

            return ((double)pixels - actualPadding) / this.maxCharWidthInPixels;
        }

        /// <summary>
        /// Converts a supplied number of pixels to a width in OpenXML
        /// </summary>
        /// <param name="pixels">The number of pixels</param>
        /// <returns>A width in OpenXML</returns>
        public double PixelsToOpenXmlWidth(int pixels)
        {
            return (double)pixels / this.maxCharWidthInPixels;
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Calculates the maximum width, in pixels, of the numeric digits 0 to 9
        /// </summary>
        /// <param name="fontFamily">The font family</param>
        /// <param name="size">The size of the font</param>
        /// <returns>The width of the widest digit in Pixels</returns>
        private double CalculateMaxDigitWidth(string fontFamily, float size)
        {
            // After a load of trawling the internet, best method is to use MeasureCharacterRanges
            var stringFont = new Font(fontFamily, size);

            // Measures the character starting at index 0, length = 1 - For some reason GenericTypographic gives the closest pixel measure
            CharacterRange[] characterRanges = { new CharacterRange(0, 1) };
            StringFormat stringFormat = new StringFormat(StringFormat.GenericTypographic);
            stringFormat.SetMeasurableCharacterRanges(characterRanges);

            double maxDigitWidth = 0.0f;

            // I just need a Graphics object. Any reasonable bitmap size will do.
            using (Graphics g = Graphics.FromImage(new Bitmap(200, 200)))
            {
                for (int i = 0; i < 10; i++)
                {
                    Region[] stringRegions = g.MeasureCharacterRanges(i.ToString(), stringFont, Rectangle.Empty, stringFormat);
                    RectangleF measureRect = stringRegions[0].GetBounds(g);

                    double widthOfDigit = (double)measureRect.Width;
                    if (widthOfDigit > maxDigitWidth)
                    {
                        maxDigitWidth = widthOfDigit;
                    }
                }
            }

            return maxDigitWidth;
        }

        #endregion Private Helpers
    }
}

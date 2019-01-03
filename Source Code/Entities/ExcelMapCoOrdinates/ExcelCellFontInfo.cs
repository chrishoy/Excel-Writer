namespace ExcelWriter
{
    using System;
    using System.Windows;
    using System.Windows.Media;

    /// <summary>
    /// Maintains information about a cell font when written to Excel.
    /// </summary>
    internal class ExcelCellFontInfo : IEquatable<ExcelCellFontInfo>, ICloneable
    {
        #region Private Fields

        private string fontFamily;
        private double fontSize;
        private FontStyle fontStyle;
        private FontWeight fontWeight;
        private Color fontColour;
        private bool fontUnderlined;

        #endregion Private Fields

        #region Public Properties

        public string FontFamily
        {
            get { return this.fontFamily; }
            set { this.fontFamily = value; }
        }

        public double FontSize
        {
            get { return this.fontSize; }
            set { this.fontSize = value; }
        }

        public FontStyle FontStyle
        {
            get { return this.fontStyle; }
            set { this.fontStyle = value; }
        }

        public FontWeight FontWeight
        {
            get { return this.fontWeight; }
            set { this.fontWeight = value; }
        }

        public Color FontColour
        {
            get { return this.fontColour; }
            set { this.fontColour = value; }
        }

        public bool FontUnderlined
        {
            get { return this.fontUnderlined; }
            set { this.fontUnderlined = value; }
        }

        #endregion Public Properties

        /// <summary>
        /// 
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(ExcelCellFontInfo other)
        {
            if (other == null) return false;

            if (!string.Equals(this.FontFamily, other.FontFamily)) return false;
            if (!double.Equals(this.FontSize, other.FontSize)) return false;
            if (!FontStyle.Equals(this.FontStyle, other.FontStyle)) return false;
            if (!System.Windows.FontWeight.Equals(this.FontWeight, other.FontWeight)) return false;
            if (!Color.Equals(this.FontColour, other.FontColour)) return false;
            if (!bool.Equals(this.FontUnderlined, other.FontUnderlined)) return false;

            return true;
        }

        public static bool Equals(ExcelCellFontInfo a, ExcelCellFontInfo b)
        {
            if (a == null && b == null) return true;    // Both null
            if (a == null && b != null) return false;   // a null, b not null
            if (b == null) return false;                // a not null, b null
            return a.Equals(b);
        }

        /// <summary>
        /// Returns a string representation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            string fontFamily = this.fontFamily == null ? string.Empty : this.FontFamily.ToString();
            return string.Format("FontInfo:{0},{1}pt,Style={2},Weight={3}", fontFamily, this.fontSize, this.fontStyle, this.fontWeight);
        }

        public object Clone()
        {
            return new ExcelCellFontInfo
            {
                FontFamily = this.FontFamily,
                FontSize = this.FontSize,
                FontStyle = this.FontStyle,
                FontWeight = this.FontWeight,
                FontUnderlined = this.FontUnderlined,
                FontColour = this.FontColour,
            };
        }
    }
}

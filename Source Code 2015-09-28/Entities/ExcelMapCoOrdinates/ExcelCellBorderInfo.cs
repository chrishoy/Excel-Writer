using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Windows;

namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about a cell's border that is to be written into an excel worksheet
    /// </summary>
    internal class ExcelCellBorderInfo : IEquatable<ExcelCellBorderInfo>, ICloneable
    {
        #region Private Fields

        private Color? colourLeft;
        private Color? colourTop;
        private Color? colourRight;
        private Color? colourBottom;
        private double widthLeft;
        private double widthTop;
        private double widthRight;
        private double widthBottom;

        #endregion Private Fields

        #region Public Properties

        public Color? ColourLeft
        {
            get { return this.colourLeft; }
            set { this.colourLeft = value; }
        }

        public Color? ColourTop
        {
            get { return this.colourTop; }
            set { this.colourTop = value; }
        }

        public Color? ColourRight
        {
            get { return this.colourRight; }
            set { this.colourRight = value; }
        }

        public Color? ColourBottom
        {
            get { return this.colourBottom; }
            set { this.colourBottom = value; }
        }

        public double WidthLeft
        {
            get { return this.widthLeft; }
            set { this.widthLeft = value; }
        }

        public double WidthTop
        {
            get { return this.widthTop; }
            set { this.widthTop = value; }
        }

        public double WidthRight
        {
            get { return this.widthRight; }
            set { this.widthRight = value; }
        }

        public double WidthBottom
        {
            get { return this.widthBottom; }
            set { this.widthBottom = value; }
        }

        /// <summary>
        /// Returns true if this info represents a visible border
        /// </summary>
        public bool HasBorder
        {
            get
            {
                return this.HasLeftBorder || this.HasTopBorder || this.HasRightBorder || this.HasBottomBorder;
            }
        }

        public bool HasLeftBorder
        {
            get { return this.widthLeft > 0 && HasColour(this.colourLeft); }
        }

        public bool HasTopBorder
        {
            get { return this.widthTop > 0 && HasColour(this.colourTop); }
        }

        public bool HasRightBorder
        {
            get { return this.widthRight > 0 && HasColour(this.colourRight); }
        }

        public bool HasBottomBorder
        {
            get { return this.widthBottom > 0 && HasColour(this.colourBottom); }
        }


        #endregion Public Properties

        private static bool HasColour(Color? color)
        {
            return color != null && color.HasValue && color.Value != Colors.Transparent;
        }

        #region Construction

        /// <summary>
        /// Ctor. Initialises an empty instance of a <see cref="ExcelCellBorderInfo"/>
        /// </summary>
        public ExcelCellBorderInfo()
        {
        }

        #endregion Construction

        public bool Equals(ExcelCellBorderInfo other)
        {
            // If both result and other have no border, return true.
            if (!this.HasBorder && !other.HasBorder) return true;
            if (!this.HasBorder && other == null) return true;
            if (other == null) return false;

            // Compare border widths, taking into account that a colour can be transparent or not set
            double thisWidthLeft = HasColour(this.ColourLeft) ? this.WidthLeft : 0;
            double otherWidthLeft = HasColour(other.ColourLeft) ? other.WidthLeft : 0;
            if (!double.Equals(thisWidthLeft, otherWidthLeft)) return false;
            if (thisWidthLeft > 0)
            {
                if (!Color.Equals(this.ColourLeft, other.ColourLeft)) return false;
            }

            double thisWidthTop = HasColour(this.ColourTop) ? this.WidthTop : 0;
            double otherWidthTop = HasColour(other.ColourTop) ? other.WidthTop : 0;
            if (!double.Equals(thisWidthTop, otherWidthTop)) return false;
            if (thisWidthTop > 0)
            {
                if (!Color.Equals(this.ColourTop, other.ColourTop)) return false;
            }

            double thisWidthRight = HasColour(this.ColourRight) ? this.WidthRight : 0;
            double otherWidthRight = HasColour(other.ColourRight) ? other.WidthRight : 0;
            if (!double.Equals(thisWidthRight, otherWidthRight)) return false;
            if (thisWidthRight > 0)
            {
                if (!Color.Equals(this.ColourRight, other.ColourRight)) return false;
            }

            double thisWidthBottom = HasColour(this.ColourBottom) ? this.WidthBottom : 0;
            double otherWidthBottom = HasColour(other.ColourBottom) ? other.WidthBottom : 0;
            if (!double.Equals(thisWidthBottom, otherWidthBottom)) return false;
            if (thisWidthBottom > 0)
            {
                if (!Color.Equals(this.ColourBottom, other.ColourBottom)) return false;
            }

            return true;
        }

        public static bool Equals(ExcelCellBorderInfo a, ExcelCellBorderInfo b)
        {
            if (a == null && b == null) return true;    // Both null
            if (a == null && b != null) return false;   // a null, b not null
            if (b == null) return false;                // a not null, b null
            return a.Equals(b);
        }


        public object Clone()
        {
            return new ExcelCellBorderInfo
            {
                ColourLeft = this.ColourLeft,
                ColourTop = this.ColourTop,
                ColourRight = this.ColourRight,
                ColourBottom = this.ColourBottom,

                WidthLeft = this.WidthLeft,
                WidthTop = this.WidthTop,
                WidthRight = this.WidthRight,
                WidthBottom = this.WidthBottom,
            };
        }

        /// <summary>
        /// Returns a string representation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            return string.Format("{0}:[{1},{2},{3},{4}]", base.ToString(), this.HasLeftBorder, this.HasTopBorder, this.HasRightBorder, this.HasBottomBorder);
        }
    }
}

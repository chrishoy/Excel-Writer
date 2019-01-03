using System;
using System.Windows;
using System.Windows.Media;

namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about a cell content alignment when written to Excel.
    /// </summary>
    internal class ExcelCellAlignmentInfo : IEquatable<ExcelCellAlignmentInfo>, ICloneable
    {
        #region Private Fields

        private double leftMargin;
        private TextAlignment textAlignment;
        private VerticalAlignment verticalAlignment;
        private TextWrapping textWrapping;
        private double textRotationAngle;

        #endregion Private Fields

        #region Public Properties

        public double LeftMargin
        {
            get { return this.leftMargin; }
            set { this.leftMargin = value; }
        }

        public TextAlignment TextAlignment
        {
            get { return this.textAlignment; }
            set { this.textAlignment = value; }
        }

        public VerticalAlignment VerticalAlignment
        {
            get { return this.verticalAlignment; }
            set { this.verticalAlignment = value; }
        }

        public TextWrapping TextWrapping
        {
            get { return this.textWrapping; }
            set { this.textWrapping = value; }
        }

        public double TextRotationAngle
        {
            get { return this.textRotationAngle; }
            set { this.textRotationAngle = value; }
        }

        #endregion Public Properties

        public bool Equals(ExcelCellAlignmentInfo other)
        {
            if (other == null) return false;

            if (!System.Windows.TextAlignment.Equals(this.TextAlignment, other.TextAlignment)) return false;
            if (!double.Equals(this.TextRotationAngle, other.TextRotationAngle)) return false;
            if (!System.Windows.TextWrapping.Equals(this.TextWrapping, other.TextWrapping)) return false;
            if (!double.Equals(this.LeftMargin, other.LeftMargin)) return false;
            if (!System.Windows.VerticalAlignment.Equals(this.VerticalAlignment, other.VerticalAlignment)) return false;

            return true;
        }

        public static bool Equals(ExcelCellAlignmentInfo a, ExcelCellAlignmentInfo b)
        {
            if (a == null && b == null) return true;    // Both null
            if (a == null && b != null) return false;   // a null, b not null
            if (b == null) return false;                // a not null, b null
            return a.Equals(b);
        }

        public object Clone()
        {
            return new ExcelCellAlignmentInfo
            {
                TextWrapping = this.TextWrapping,
                TextAlignment = this.TextAlignment,
                TextRotationAngle = this.TextRotationAngle,
                LeftMargin = this.LeftMargin,
                VerticalAlignment = this.VerticalAlignment,
            };
        }
    }
}

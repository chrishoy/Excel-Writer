using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Windows;

namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about the style of a cell that is to be written into an excel worksheet
    /// </summary>
    internal class ExcelCellStyleInfo : IEquatable<ExcelCellStyleInfo>, ICloneable
    {
        #region Private Fields

        private Color? fillColour;
        private ExcelCellBorderInfo borderInfo;
        private ExcelCellFontInfo fontInfo;
        private ExcelCellAlignmentInfo alignmentInfo;
        private string numberFormat;

        private bool hasCellInfo;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor. Initialises and updates a new instance of a <see cref="ExcelCellStyleInfo"/> with supplied cell information
        /// </summary>
        /// <param name="Row"></param>
        public ExcelCellStyleInfo(ExcelMapCoOrdinate cell)
        {
            if (cell == null) throw new ArgumentNullException("cell");

            if (cell is ExcelMapCoOrdinateCell)
            {
                this.ApplyCellStyles(((ExcelMapCoOrdinateCell)cell).Styles);
            }
            if (cell is ExcelMapCoOrdinateContainer)
            {
                if (cell.Styles != null)
                {
                    foreach (var style in cell.Styles)
                    {
                        this.ApplyContainerStyles(((ExcelMapCoOrdinateContainer)cell).Styles);
                    }
                }
            }
        }

        /// <summary>
        /// Ctor. Initialises an empty instance of a <see cref="ExcelCellStyleInfo"/>
        /// </summary>
        public ExcelCellStyleInfo()
        {
        }

        /// <summary>
        /// Creates a copy of original.
        /// </summary>
        /// <param name="source"></param>
        private ExcelCellStyleInfo(ExcelCellStyleInfo source)
        {
            this.fillColour = source.FillColour;
            this.numberFormat = source.NumberFormat;
            this.hasCellInfo = source.HasCellInfo;

            this.fontInfo = source.FontInfo == null ? null : (ExcelCellFontInfo)source.FontInfo.Clone();
            this.borderInfo = source.BorderInfo == null ? null : (ExcelCellBorderInfo)source.BorderInfo.Clone();
            this.alignmentInfo = source.AlignmentInfo == null ? null : (ExcelCellAlignmentInfo)source.AlignmentInfo.Clone();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Returns true if there is any Cell information used to format Cell in an excel worksheet
        /// </summary>
        public bool HasCellInfo
        {
            get { return this.hasCellInfo; }
            set { this.hasCellInfo = value; }
        }

        /// <summary>
        /// Width and colours of a borders around this cell when written to Excel.
        /// </summary>
        public ExcelCellBorderInfo BorderInfo
        {
            get { return this.borderInfo; }
            set { this.borderInfo = value; }
        }

        /// <summary>
        /// Colour that will be used to fill the Excel cell
        /// </summary>
        public Color? FillColour
        {
            get { return this.fillColour; }
            set { this.fillColour = value; }
        }

        /// <summary>
        /// Format string to be applied to this cell when written to Excel.
        /// </summary>
        public string NumberFormat
        {
            get { return this.numberFormat; }
            set { this.numberFormat = value; }
        }

        /// <summary>
        /// Gets information about the font to be applied in the Excel cell
        /// </summary>
        public ExcelCellFontInfo FontInfo
        {
            get { return this.fontInfo; }
            set { this.fontInfo = value; }
        }

        /// <summary>
        /// Gets information about the cell alignment to be applied in the Excel cell
        /// </summary>
        public ExcelCellAlignmentInfo AlignmentInfo
        {
            get { return this.alignmentInfo; }
            set { this.alignmentInfo = value; }
        }

        #endregion Public Properties

        #region Methods

        /// <summary>
        /// Compares and updates this information based on supplied cell
        /// </summary>
        /// <param name="row"></param>
        public void ApplyCellStyles(IEnumerable<StyleBase> styles)
        {
            if (styles == null) return; // No information to update

            foreach (var style in styles)
            {
                this.UpdateCellStyle(style);
            }
        }

        /// <summary>
        /// Compares and updates this information based on supplied cell
        /// </summary>
        /// <param name="row"></param>
        public void ApplyContainerStyles(IEnumerable<StyleBase> styles)
        {
            if (styles == null) return; // No information to update

            foreach (var style in styles)
            {
                // Update fill colour (if not already set)
                this.UpdateFillColour(style);

                // Don't update borders as this is conditional on location of current cell within container.

                // Update font (if not already set)
                this.UpdateFontInfo(style);

                // Update alignment (if not already set)
                this.UpdateAlignmentInfo(style);

                // Update the excel number format (if not already set)
                this.UpdateNumberFormat(style);
            }
        }

        /// <summary>
        /// Updates cell formatting based on supplied co-ordinate and the co-ordinate position within the container.
        /// Formatting may be influenced by the current position within container.
        /// </summary>
        /// <param name="container">The container this cell is within</param>
        /// <param name="fromColumnIndex">The excel 'from' column</param>
        /// <param name="fromRowIndex">The excel 'from' row</param>
        /// <param name="toColumnIndex">The excel 'to' column</param>
        /// <param name="toRowIndex">The excel 'to' row</param>
        public void Update(ExcelMapCoOrdinateContainer container, int fromColumnIndex, int fromRowIndex, int toColumnIndex, int toRowIndex)
        {
            if (container == null) throw new ArgumentNullException("container");

            if (container.Styles != null)
            {
                // Update border thickness (if there is one)
                ExcelCellBorderInfo conditionalBorderInfo = CreateConditionalBorderInfo(container, fromColumnIndex, fromRowIndex, toColumnIndex, toRowIndex);
                this.UpdateBorderInfo(conditionalBorderInfo);

                foreach (var style in container.Styles)
                {
                    // Update fill colour (if not already set)
                    this.UpdateFillColour(style);

                    // Update font (if not already set)
                    this.UpdateFontInfo(style);

                    // Update alignment (if not already set)
                    this.UpdateAlignmentInfo(style);

                    // Update the excel number format (if not already set)
                    this.UpdateNumberFormat(style);
                }
            }
        }

        /// <summary>
        /// Updates this <see cref="ExcelCellStyleInfo"/> with <see cref="StyleBase">Map Style</see> properties.
        /// </summary>
        /// <param name="exportStyle"></param>
        private void UpdateCellStyle(StyleBase style)
        {
            // Update fill colour (if not already set)
            this.UpdateFillColour(style);

            // Update border (this is layered)
            this.UpdateBorderInfo(style);

            // Update font (if not already set)
            this.UpdateFontInfo(style);

            // Update alignment (if not already set)
            this.UpdateAlignmentInfo(style);

            // Update the excel number format (if not already set)
            this.UpdateNumberFormat(style);
        }

        /// <summary>
        /// Creates and returns a conditional border thickness which is dependent on where we are within the container.
        /// </summary>
        /// <param name="container">The container this cell is within</param>
        /// <param name="fromColumnIndex">The excel 'from' column</param>
        /// <param name="fromRowIndex">The excel 'from' row</param>
        /// <param name="toColumnIndex">The excel 'to' column</param>
        /// <param name="toRowIndex">The excel 'to' row</param>
        /// <returns></returns>
        private static ExcelCellBorderInfo CreateConditionalBorderInfo(ExcelMapCoOrdinateContainer container,
                                                                       int fromColumnIndex, int fromRowIndex, int toColumnIndex, int toRowIndex)
        {
            var newBorderInfo = new ExcelCellBorderInfo();
            if (container.Styles != null)
            {
                foreach (var style in container.Styles)
                {
                    // Update border thickness (if there is one either conditionally on location or unconditionally)
                    if (style.HasAnyBorder())
                    {
                        // Only apply if currently left-most
                        if (container.ExcelColumnStart == fromColumnIndex && style.BorderThickness.Value.Left > 0)
                        {
                            newBorderInfo.WidthLeft = style.BorderThickness.Value.Left;
                            newBorderInfo.ColourLeft = style.BorderColour.Value;
                        }

                        // Only apply if currently top-most
                        if (container.ExcelRowStart == fromRowIndex && style.BorderThickness.Value.Top > 0)
                        {
                            newBorderInfo.WidthTop = style.BorderThickness.Value.Top;
                            newBorderInfo.ColourTop = style.BorderColour.Value;
                        }

                        // Only apply if currently right-most
                        uint endColumnIndex = container.GetEndColumnIndex();
                        if (endColumnIndex == toColumnIndex && style.BorderThickness.Value.Right > 0)
                        {
                            newBorderInfo.WidthRight = style.BorderThickness.Value.Right;
                            newBorderInfo.ColourRight = style.BorderColour.Value;
                        }

                        // Only apply if currently bottom-most
                        uint endRowIndex = container.GetEndRowIndex();
                        if (endRowIndex == toRowIndex && style.BorderThickness.Value.Bottom > 0)
                        {
                            newBorderInfo.WidthBottom = style.BorderThickness.Value.Bottom;
                            newBorderInfo.ColourBottom = style.BorderColour.Value;
                        }
                    }
                }
            }
            return newBorderInfo;
        }

        /// <summary>
        /// Compares and updates the fill colour for the cell
        /// </summary>
        private void UpdateFillColour(StyleBase mapStyle)
        {
            if (mapStyle != null)
            {
                //Only applies colour if exists, is not transparent and is not already set.
                if (this.fillColour == null || this.fillColour.HasValue == false || this.fillColour.Value == Colors.Transparent)
                {
                    if (mapStyle.BackgroundColour != null && mapStyle.BackgroundColour.HasValue && mapStyle.BackgroundColour.Value != Colors.Transparent)
                    {
                        this.fillColour = mapStyle.BackgroundColour;
                        this.hasCellInfo = true;
                    }
                }
            }
        }

        /// <summary>
        /// Compares and updates the number format for the cell
        /// </summary>
        private void UpdateNumberFormat(StyleBase mapStyle)
        {
            var typedMapStyle = mapStyle as CellStyle;

            if (typedMapStyle != null)
            {
                //Only applies colour if exists, is not transparent and is not already set.
                if (string.IsNullOrEmpty(this.numberFormat) && !string.IsNullOrEmpty(typedMapStyle.ExcelFormat))
                {
                    this.numberFormat = typedMapStyle.ExcelFormat;
                    this.hasCellInfo = true;
                }
            }
        }

        /// <summary>
        /// Compares and updates border thickness for the cell
        /// </summary>
        /// <param name="rowHeight"></param>
        private void UpdateBorderInfo(StyleBase mapStyle)
        {
            // Anything to apply?
            if (mapStyle != null && mapStyle.HasAnyBorder())
            {
                if (this.borderInfo == null)
                {
                    // Clone the border and apply to the cell
                    this.borderInfo = CreateBorderInfo(mapStyle);
                    this.hasCellInfo = true;
                }
                else
                {
                    // Update the border with layered information
                    this.borderInfo.UpdateBorder(mapStyle);
                }
            }
        }

        /// <summary>
        /// Compares and updates border thickness for the cell.<br/>
        /// </summary>
        /// <param name="rowHeight"></param>
        private void UpdateBorderInfo(ExcelCellBorderInfo borderInfo)
        {
            // Anything to apply?
            if (borderInfo.HasBorder)
            {
                // Any current border?
                if (this.borderInfo == null)
                {
                    // No current border set, but we have a border to set - Apply
                    this.borderInfo = borderInfo;
                    this.hasCellInfo = true;
                }
                else
                {
                    // Update the border with layered information
                    this.borderInfo.UpdateBorder(borderInfo);
                }
            }
        }

        /// <summary>
        /// Creates a <see cref="ExcelCellBorderInfo"/> from the border colour and thickness specified in a <see cref="StyleBase"/> derived class instance
        /// which has uniform border thickness and border colour.
        /// NB! This is intended to raise an error if either border colour or thickness is null/not specified.
        /// </summary>
        /// <param name="mapStyle"></param>
        /// <returns></returns>
        private static ExcelCellBorderInfo CreateBorderInfo(StyleBase mapStyle)
        {
            var borderInfo = new ExcelCellBorderInfo();

            borderInfo.ColourLeft = mapStyle.BorderColour.Value;
            borderInfo.ColourTop = mapStyle.BorderColour.Value;
            borderInfo.ColourRight = mapStyle.BorderColour.Value;
            borderInfo.ColourBottom = mapStyle.BorderColour.Value;

            borderInfo.WidthLeft = mapStyle.BorderThickness.Value.Left;
            borderInfo.WidthTop = mapStyle.BorderThickness.Value.Top;
            borderInfo.WidthRight = mapStyle.BorderThickness.Value.Right;
            borderInfo.WidthBottom = mapStyle.BorderThickness.Value.Bottom;

            return borderInfo;
        }

        /// <summary>
        /// Compares and updates font information for the cell.
        /// </summary>
        /// <param name="font"></param>
        private void UpdateFontInfo(StyleBase mapStyle)
        {
            // Only applies to CellMapStyles
            CellStyle typedMapStyle = mapStyle as CellStyle;

            if (typedMapStyle != null && HasAnyFontInfo(typedMapStyle))
            { 
                if (this.fontInfo == null) this.fontInfo = new ExcelCellFontInfo();

                // Update any font-related properties that are non-null/empty
                if (!string.IsNullOrEmpty(typedMapStyle.FontFamily))
                {
                    this.fontInfo.FontFamily = typedMapStyle.FontFamily;
                }

                if (typedMapStyle.FontSize.HasValue) this.fontInfo.FontSize = typedMapStyle.FontSize.Value;
                if (typedMapStyle.FontWeight.HasValue) this.fontInfo.FontWeight = typedMapStyle.FontWeight.Value;
                if (typedMapStyle.FontUnderlined.HasValue) this.fontInfo.FontUnderlined = typedMapStyle.FontUnderlined.Value;

                if (typedMapStyle.InternalFontColour.HasValue) this.fontInfo.FontColour = typedMapStyle.InternalFontColour.Value;

                this.hasCellInfo = true;
            }
        }

        private static bool HasAnyFontInfo(CellStyle mapStyle)
        {
            if (mapStyle != null)
            {
                if (!string.IsNullOrEmpty(mapStyle.FontFamily)) return true;
                if (mapStyle.FontSize.HasValue) return true;
                if (mapStyle.FontWeight.HasValue) return true;
                if (mapStyle.FontUnderlined.HasValue) return true;
                if (mapStyle.InternalFontColour.HasValue) return true;
            }
            return false;
        }

        /// <summary>
        /// Compares and updates alignment information for the cell.
        /// </summary>
        /// <param name="font"></param>
        private void UpdateAlignmentInfo(StyleBase mapStyle)
        {
            // Only applies to CellMapStyles
            CellStyle typedMapStyle = mapStyle as CellStyle;

            if (typedMapStyle != null && HasAnyAlignmentInfo(typedMapStyle))
            {
                if (this.alignmentInfo == null) this.alignmentInfo = new ExcelCellAlignmentInfo();

                // Update any allignment-related properties that are non-null/empty
                if (typedMapStyle.TextAlignment.HasValue) this.alignmentInfo.TextAlignment = typedMapStyle.TextAlignment.Value;
                if (typedMapStyle.RotationAngle.HasValue && typedMapStyle.RotationAngle.Value != 0) this.alignmentInfo.TextRotationAngle = typedMapStyle.RotationAngle.Value;
                if (typedMapStyle.TextWrapping.HasValue) this.alignmentInfo.TextWrapping = typedMapStyle.TextWrapping.Value;
                if (typedMapStyle.Indentation.HasValue) this.alignmentInfo.LeftMargin = typedMapStyle.Indentation.Value * 2; // NB! CH-Can we remove this /2... It's there because we / 2 when applying to cell
                if (typedMapStyle.VerticalAlignment.HasValue) this.alignmentInfo.VerticalAlignment = typedMapStyle.VerticalAlignment.Value;
                this.hasCellInfo = true;
            }
        }

        private static bool HasAnyAlignmentInfo(CellStyle mapStyle)
        {
            if (mapStyle != null)
            {
                if (mapStyle.TextWrapping.HasValue) return true;
                if (mapStyle.TextAlignment.HasValue) return true;
                if (mapStyle.RotationAngle.HasValue) return true;
                if (mapStyle.Indentation.HasValue) return true;
                if (mapStyle.VerticalAlignment.HasValue) return true;
            }
            return false;
        }

        #endregion Methods

        #region IEquatable Members

        /// <summary>
        /// Returns true if this = other
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(ExcelCellStyleInfo other)
        {
            if (other == null && !this.HasCellInfo) return true;
            if (other == null) return false;

            if (!Nullable<Color>.Equals(this.FillColour, other.FillColour)) return false;
            if (!string.Equals(this.NumberFormat, other.NumberFormat)) return false;
            if (!ExcelCellFontInfo.Equals(this.FontInfo, other.FontInfo)) return false;
            if (!ExcelCellBorderInfo.Equals(this.BorderInfo, other.BorderInfo)) return false;
            if (!ExcelCellAlignmentInfo.Equals(this.AlignmentInfo, other.AlignmentInfo)) return false;

            return true;
        }

        /// <summary>
        /// Returns true if a = b
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public static bool Equals(ExcelCellStyleInfo a, ExcelCellStyleInfo b)
        {
            if (a == null && b == null) return true;
            if (a == null && b != null) return false;
            return a.Equals(b);
        }

        #endregion

        #region ICloneable Members

        /// <summary>
        /// Create a new instance of this object
        /// </summary>
        /// <returns></returns>
        public object Clone()
        {
            return new ExcelCellStyleInfo(this);
        }

        #endregion

    }
}

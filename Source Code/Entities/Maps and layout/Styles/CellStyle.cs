namespace ExcelWriter
{
    using System;
    using System.Windows.Media;
    using System.Windows;

    /// <summary>
    /// Style attributes which can be specifically used with <see cref="Cell"/>s.
    /// </summary>
    public class CellStyle : StyleBase
    {
        private double? fontSize; // = 10;
        private string fontFamily; // = "Calibri";
        private bool? fontUnderlined;
        private FontWeight? fontWeight;
        private TextAlignment? textAlignment; // = TextAlignment.Center;
        private TextWrapping? textWrapping; //= TextWrapping.NoWrap;
        private VerticalAlignment? verticalAlignment; // = VerticalAlignment.Bottom;
        private HorizontalAlignment? horizontalAlignment; // = HorizontalAlignment.Center;
        private string excelFormat;

        public double? rotationAngle;
        public int? indentation;

        // Bindable
        private object fontColour; // = Colors.Black;

        #region Public properties

        /// <summary>
        /// Get/set Excel formatting string used for cells in this map.
        /// </summary>
        public string ExcelFormat
        {
            get { return this.excelFormat; }
            set { this.excelFormat = value; }
        }

        /// <summary>
        /// Get/set the font/text colour for the cells in this map.
        /// </summary>
        public object FontColour
        {
            get { return BindingContainer.EvaluateIfRequired(this.fontColour, this.DataContext); }
            set { this.fontColour = BindingContainer.CreateIfRequired(value); }
        }

        /// <summary>
        /// Get/set size of font for the cells in this map.
        /// </summary>
        public double? FontSize
        {
            get { return this.fontSize; }
            set { this.fontSize = value; }
        }

        /// <summary>
        /// Get/set whether font is underlined for the cells in this map.
        /// </summary>
        public bool? FontUnderlined
        {
            get { return this.fontUnderlined; }
            set { this.fontUnderlined = value; }
        }

        /// <summary>
        /// Get/set name of font family for the cells in this map.
        /// </summary>
        public string FontFamily
        {
            get { return this.fontFamily; }
            set { this.fontFamily = value; }
        }

        /// <summary>
        /// Get/set font weight for the cells in this map.
        /// </summary>
        public FontWeight? FontWeight
        {
            get { return this.fontWeight; }
            set { this.fontWeight = value; }
        }

        /// <summary>
        /// Get/set text alignment for the cells in this map.
        /// </summary>
        public TextAlignment? TextAlignment
        {
            get { return this.textAlignment; }
            set { this.textAlignment = value; }
        }

        /// <summary>
        /// Get/set text wrapping within cells in this map.
        /// </summary>
        public TextWrapping? TextWrapping
        {
            get { return this.textWrapping; }
            set { this.textWrapping = value; }
        }

        /// <summary>
        /// Get/set vertical alignment of the content of the cells in this map.
        /// </summary>
        public VerticalAlignment? VerticalAlignment
        {
            get { return this.verticalAlignment; }
            set { this.verticalAlignment = value; }
        }

        /// <summary>
        /// Get/set horizontal alignment of the content of the cells in this map.
        /// </summary>
        public HorizontalAlignment? HorizontalAlignment
        {
            get { return this.horizontalAlignment; }
            set { this.horizontalAlignment = value; }
        }

        /// <summary>
        /// Get/set the rotation angle of the content of the cells in this map.
        /// </summary>
        public double? RotationAngle
        {
            get { return this.rotationAngle; }
            set { this.rotationAngle = value; }
        }

        /// <summary>
        /// Get/set the indentation of the content of the cells in this map.
        /// </summary>
        public int? Indentation
        {
            get { return this.indentation; }
            set { this.indentation = value; }
        }

        #endregion

        #region Internal properties

        /// <summary>
        /// Returns a Color? from the bindable FontColour property
        /// </summary>
        internal Color? InternalFontColour
        {
            get
            {
                // if nothing return nothing
                if (this.FontColour == null)
                {
                    return null;
                }

                try
                {
                    // use the inbuilt ColorConverter is the FontColour is a string
                    // should handle SystemColors such as 'White', 'Black' and hex strings.
                    if (this.FontColour is string)
                    {
                        return (Color)ColorConverter.ConvertFromString((string)this.FontColour);
                    }
                    // if the object is a Color then safe is cast
                    else if (this.FontColour is Color)
                    {
                        return (Color)this.FontColour;
                    }
                    // otherwise return null
                    return null;
                }
                catch (Exception ex)
                {
                    // any exception are wrapped and rethrown, might be best to just swallow and log
                    throw new MetadataException("Only Color can be bound to FontColour", ex);
                }
            }
        }

        #endregion

        #region IClonable members

        /// <summary>
        /// Create and return a copy of this instance.
        /// </summary>
        /// <returns></returns>
        public override object Clone()
        {
            return new CellStyle
            {
                BackgroundColour = this.BackgroundColour,
                BasedOnKey = this.BasedOnKey,
                BorderColour = this.BorderColour,
                BorderThickness = this.BorderThickness,
                Indentation = this.Indentation,
                Key = this.Key,
                FontColour = this.FontColour,
                FontSize = this.FontSize,
                FontFamily = this.FontFamily,
                FontWeight = this.FontWeight,
                FontUnderlined = this.FontUnderlined,
                TextAlignment = this.TextAlignment,
                TextWrapping = this.TextWrapping,
                VerticalAlignment = this.VerticalAlignment,
                HorizontalAlignment = this.HorizontalAlignment,
                ExcelFormat = this.ExcelFormat,
                RotationAngle = this.RotationAngle,
            };
        }

        #endregion IClonable members

        /// <summary>
        /// Creates a new style which is based on an existing style, merging over current values.
        /// </summary>
        /// <param name="styleToMerge">Style which is to be merged over the base style</param>
        /// <param name="basedOnStyle">Style on which the new style is based</param>
        /// <returns></returns>
        public static CellStyle CreateMergedStyle(StyleBase basedOnStyle, CellStyle styleToMerge)
        {
            // First clone
            var newStyle = (CellStyle)styleToMerge;

            // Then override
            newStyle.BackgroundColour = styleToMerge.BackgroundColour.HasValue ? styleToMerge.BackgroundColour : basedOnStyle.BackgroundColour;
            newStyle.BorderColour = styleToMerge.BorderColour.HasValue ? styleToMerge.BorderColour : basedOnStyle.BorderColour;
            newStyle.BorderThickness = styleToMerge.BorderThickness.HasValue ? styleToMerge.BorderThickness : basedOnStyle.BorderThickness;

            // If supplied is style is CellStyle then apply these attributes also
            if (basedOnStyle is CellStyle)
            {
                var typedValueToMerge = basedOnStyle as CellStyle;

                // using new internal readonly FontColour property which gets a typed verison of the bindable Color? FontColour
                newStyle.FontColour = styleToMerge.InternalFontColour.HasValue ? styleToMerge.InternalFontColour : typedValueToMerge.InternalFontColour;

                newStyle.ExcelFormat = !string.IsNullOrEmpty(styleToMerge.ExcelFormat) ? styleToMerge.ExcelFormat : typedValueToMerge.ExcelFormat;
                newStyle.FontFamily = !string.IsNullOrEmpty(styleToMerge.FontFamily) ? styleToMerge.FontFamily : typedValueToMerge.FontFamily;
                newStyle.FontSize = styleToMerge.FontSize.HasValue ? styleToMerge.FontSize : typedValueToMerge.FontSize;
                newStyle.FontWeight = styleToMerge.FontWeight.HasValue ? styleToMerge.FontWeight : typedValueToMerge.FontWeight;
                newStyle.FontUnderlined = styleToMerge.FontUnderlined.HasValue ? styleToMerge.FontUnderlined : typedValueToMerge.FontUnderlined;
                newStyle.HorizontalAlignment = styleToMerge.HorizontalAlignment.HasValue ? styleToMerge.HorizontalAlignment : typedValueToMerge.HorizontalAlignment;
                newStyle.RotationAngle = styleToMerge.RotationAngle.HasValue ? styleToMerge.RotationAngle : typedValueToMerge.RotationAngle;
                newStyle.TextAlignment = styleToMerge.TextAlignment.HasValue ? styleToMerge.TextAlignment : typedValueToMerge.TextAlignment;
                newStyle.TextWrapping = styleToMerge.TextWrapping.HasValue ? styleToMerge.TextWrapping : typedValueToMerge.TextWrapping;
                newStyle.VerticalAlignment = styleToMerge.VerticalAlignment.HasValue ? styleToMerge.VerticalAlignment : typedValueToMerge.VerticalAlignment;
            }

            return newStyle;
        }

    }
}

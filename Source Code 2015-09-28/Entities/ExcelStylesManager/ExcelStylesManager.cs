namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;

    using Constants;
    using OpenXml.Excel;
    using DocumentFormat.OpenXml;
    using System.Windows;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Manages styles
    /// </summary>
    internal class ExcelStylesManager
    {
        #region Private Fields

        private readonly Stylesheet stylesheet;
        private readonly StylesDictionary styles;
        private readonly FillsDictionary fills;
        private readonly FontsDictionary fonts;
        private readonly NumberFormatsDictionary numberFormats;
        private readonly BordersDictionary borders;
        private readonly StylesCollection sharedMapStyles;
        private readonly StylesCollection currentMapStyles;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor. Creates in instance of a <see cref="ExcelStylesManager"/> which is used to map<br/>
        /// <see cref="ExcelCellInfo">information about a cell that is to be written to Excel</see><br/>
        /// to Excel <see cref="Stylesheet"/> styles (keyed using uint)
        /// </summary>
        /// <param name="document">OpenXML workbook</param>
        public ExcelStylesManager(SpreadsheetDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");

            // Load shared map styles which are stored as a local resource
            this.sharedMapStyles = StylesCollection.Deserialize(Properties.Resources.SharedMapStyles);
            this.currentMapStyles = new StylesCollection();

            this.stylesheet = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;

            this.styles = new StylesDictionary();
            this.fills = new FillsDictionary();
            this.fonts = new FontsDictionary();
            this.numberFormats = new NumberFormatsDictionary();
            this.borders = new BordersDictionary();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets a reference to the <see cref="Stylesheet"/> of the Excel document that is being created.
        /// </summary>
        public Stylesheet Stylesheet
        {
            get { return this.stylesheet; }
        }

        public ResourceStore ResourceStore
        {
            get;
            private set;
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// 
        /// </summary>
        public void Initialise(ResourceStore resourceStore)
        {
            if (resourceStore == null)
            {
                throw new ArgumentNullException("resourceStore");
            }

            this.currentMapStyles.Clear();
            foreach (var mapStyle in this.sharedMapStyles)
            {
                // Merges the style into the current collection.
                this.currentMapStyles.MergeCloneStyle(mapStyle);
            }

            this.ResourceStore = resourceStore;

            // Merges the style into the current collection.
            foreach (var mapStyle in this.currentMapStyles)
            {
                this.ResourceStore.Add(mapStyle.Key, ResourceTypeNames.StyleBase, mapStyle, null);
            }
        }

        /// <summary>
        /// Sets the current template <see cref="ExcelStyleResources">style resources</see> that are to be used when registering styles.
        /// </summary>
        /// <param name="styleResources"></param>
        public void SetCurrentMapStyles(IEnumerable<StyleBase> styleResources)
        {
            // Clear down all current map styles and re-apply as certain styles may be overriden with local styles
            this.currentMapStyles.Clear();
            foreach (StyleBase mapStyle in this.sharedMapStyles)
            {
                // Merges the style into the current collection.
                this.currentMapStyles.MergeCloneStyle(mapStyle);
            }

            // Now add the current set.
            foreach (StyleBase mapStyle in styleResources)
            {
                this.currentMapStyles.MergeCloneStyle(mapStyle);
            }

            this.ResourceStore = new ResourceStore();

            // Merges the style into the current collection.
            foreach (var mapStyle in this.currentMapStyles)
            {
                this.ResourceStore.Add(mapStyle.Key, ResourceTypeNames.StyleBase, mapStyle, null);
            }
        }

        public StyleBase Merge(StyleBase value)
        {
            if (!string.IsNullOrEmpty(value.BasedOnKey))
            {
                // Look up mandatory existing style
                StyleBase baseStyle = this.ResourceStore.GetResourceByKey<StyleBase>(value.BasedOnKey);
                StyleBase newStyle;

                // And update with properties of the supplied style
                if (value is Style)
                {
                    if (baseStyle is Style)
                    {
                        // ExcelMapStyle based on ExcelMapStyle
                        newStyle = Style.CreateMergedStyle(baseStyle, (Style)value);
                    }
                    else if (baseStyle is CellStyle)
                    {
                        // ExcelMapStyle can't be based on a ExcelCellMapStyle
                        throw new InvalidOperationException(string.Format("Can't base ExcelMapStyle '{0}'  on ExcelCellMapStyle '{1}' (not yet anyway)", value.Key, baseStyle.Key));
                    }
                    else
                    {
                        // Whoops.. not an accounted for style
                        throw new InvalidOperationException(string.Format("Style Key='{0}' BasedOnStyle is not valid style", baseStyle.Key));
                    }

                }
                else if (value is CellStyle)
                {
                    if (baseStyle is Style)
                    {
                        // ExcelCellMapStyle based on ExcelMapStyle
                        newStyle = CellStyle.CreateMergedStyle(baseStyle, (CellStyle)value);
                    }
                    else if (baseStyle is CellStyle)
                    {
                        // Straight Copy & Merge
                        newStyle = CellStyle.CreateMergedStyle(baseStyle, (CellStyle)value);
                    }
                    else
                    {
                        // Whoops.. not an accounted for style
                        throw new InvalidOperationException(string.Format("Style Key='{0}' BasedOnStyle is not valid style", baseStyle.Key));
                    }
                }
                else
                {
                    // Whoops.. not an accounted for style
                    throw new InvalidOperationException(string.Format("Style Key='{0}' is not valid style", value.Key));
                }

                return newStyle;
            }
            return value;            
        }

        /// <summary>
        /// Get a style based on supplied <see cref="ExcelCellStyleInfo"/>.<br/>
        /// The style may have to be created in the stylesheet.
        /// </summary>
        /// <param name="cellInfo"></param>
        /// <returns>Index of the style in the stylesheet</returns>
        public uint GetOrCreateStyle(ExcelCellStyleInfo cellInfo)
        {
            return this.GetStyle(cellInfo, string.Empty);
        }

        /// <summary>
        /// Get a style based on supplied <see cref="ExcelCellInfo"/>.<br/>
        /// The style may have to be created in the stylesheet.
        /// </summary>
        /// <param name="mapStyle">Information about the style that has be be created.</param>
        /// <returns>Index of the style in the stylesheet</returns>
        public uint GetOrCreateStyle(StyleBase mapStyle)
        {
            System.Diagnostics.Debug.Assert(false, "TO BE WRITTEN");
            return 0;
        }

        /// <summary>
        /// Get a style based on supplied <see cref="ExcelCellInfo"/> and excel format code.<br/>
        /// The style may have to be created in the stylesheet.
        /// </summary>
        /// <param name="cellInfo"></param>
        /// <param name="formatCode"></param>
        /// <returns>Index of the style in the stylesheet</returns>
        public uint GetStyle(ExcelCellStyleInfo cellInfo, string formatCode)
        {
            // No cell styling information to apply, so return 0
            if (!cellInfo.HasCellInfo) return 0;

            uint styleIndex;

            var existing = this.styles.Find(cellInfo);

            if (existing.Key != null)
            {
                styleIndex = existing.Value;
            }
            else
            {
                styleIndex = CreateNewStyle(this, this.stylesheet, cellInfo, formatCode);
                this.styles.Add((ExcelCellStyleInfo)cellInfo.Clone(), styleIndex);
            }

            return styleIndex;
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Creates a new style in the excel <see cref="Stylesheet"/> based on the supplied <see cref="ExportStyle"/>.<br/>
        /// Return the index of the newly created style.
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="cellInfo"></param>
        /// <param name="formatCode"></param>
        /// <returns></returns>
        private static uint CreateNewStyle(ExcelStylesManager stylesManager, Stylesheet stylesheet, ExcelCellStyleInfo cellInfo, string formatCode)
        {
            // Create a cellFormat to which style attributes can be applied.
            var cellFormat = new CellFormat();

            // Looks up/creates a new fill colour in the stylesheet and applies it to the cellFormat
            UpdateFillColour(stylesManager, cellInfo, ref stylesheet, ref cellFormat);

            // Create and append a new font to the stylesheet and applies it to the cellFormat
            UpdateFont(stylesManager, cellInfo, ref stylesheet, ref cellFormat);

            // Create a number format in the stylesheet and apply to the cellFormat
            UpdateNumberFormat(stylesManager, cellInfo, formatCode, ref stylesheet, ref cellFormat);

            // Creates a border in the stylesheet and applies it to the cellFormat
            UpdateBorder(stylesManager, cellInfo, ref stylesheet, ref cellFormat);

            // Create an Allignment object which manages indentation, alignment and text rotation and applies it to the cellFormat.
            UpdateTextAlignmentAndRotation(cellInfo, ref cellFormat);

            uint styleIndex = stylesheet.AddCellFormat(cellFormat);
            return styleIndex;
        }

        /// <summary>
        /// Updates the supplied <see cref="CellFormat"/> with a colour index which relates to the Fill Colour.
        /// The colour is created in the excel <see cref="Stylesheet"/> if it is not already there.
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="cellInfo"></param>
        /// <returns></returns>
        private static void UpdateFillColour(ExcelStylesManager stylesManager, ExcelCellStyleInfo cellInfo, ref Stylesheet stylesheet, ref CellFormat cellFormat)
        {
            if (cellFormat == null) throw new ArgumentNullException("cellFormat");

            if (cellInfo.FillColour != null && cellInfo.FillColour.HasValue && cellInfo.FillColour.Value != System.Windows.Media.Colors.Transparent)
            {
                var fillItem = stylesManager.fills.Find(cellInfo.FillColour.Value);
                UInt32Value fillId = fillItem.Value;

                if (fillId == null)
                {
                    // Add and return the index of a fill colour (Badly named as simply returns index of existing colour if it already exists)
                    fillId = stylesheet.AddIndexedColor(cellInfo.FillColour.Value);
                    stylesManager.fills.Add(cellInfo.FillColour.Value, fillId);
                }
                cellFormat.FillId = fillId;
            }
        }

        /// <summary>
        /// Updates the supplied <see cref="CellFormat"/> with a font index which relates to the fund in the supplied <see cref="ExportStyle"/>
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="cellInfo"></param>
        /// <returns></returns>
        private static void UpdateFont(ExcelStylesManager stylesManager, ExcelCellStyleInfo cellInfo, ref Stylesheet stylesheet, ref CellFormat cellFormat)
        {
            if (cellFormat == null) throw new ArgumentNullException("cellFormat");

            // Anything to apply?
            if (cellInfo.FontInfo == null) return;

            // Look up existing...
            var fontItem = stylesManager.fonts.Find(cellInfo.FontInfo);
            UInt32Value id = fontItem.Value;

            if (id == null)
            {
                // Fonts
                Font font = new Font();

                if (cellInfo.FontInfo.FontFamily != null)
                {
                    font.FontName = new FontName()
                    {
                        Val = new StringValue(cellInfo.FontInfo.FontFamily.ToString())
                    };
                }

                font.FontSize = new FontSize()
                {
                    Val = new DoubleValue(cellInfo.FontInfo.FontSize)
                };

                if ((cellInfo.FontInfo.FontWeight != null) &&
                    (cellInfo.FontInfo.FontWeight == FontWeights.Bold))
                {
                    font.Bold = new Bold();
                }

                if (cellInfo.FontInfo.FontStyle != null)
                {
                    if (cellInfo.FontInfo.FontStyle == FontStyles.Italic)
                    {
                        font.Italic = new Italic();
                    }
                }

                font.Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
                {
                    Rgb = new HexBinaryValue(StyleTranslator.Translate(cellInfo.FontInfo.FontColour))
                };

                if (cellInfo.FontInfo.FontUnderlined)
                {
                    font.Underline = new Underline();
                }

                id = stylesheet.AddFont(font);
                stylesManager.fonts.Add(cellInfo.FontInfo, id);
            }

            cellFormat.FontId = id;
        }

        /// <summary>
        ///  Create a number format in the stylesheet and apply to the cellFormat
        /// </summary>
        /// <param name="cellInfo"></param>
        /// <param name="formatCode"></param>
        /// <param name="stylesheet"></param>
        /// <param name="cellFormat"></param>
        private static void UpdateNumberFormat(ExcelStylesManager stylesManager, ExcelCellStyleInfo cellInfo, string formatCode, ref Stylesheet stylesheet, ref CellFormat cellFormat)
        {
            if (cellFormat == null) throw new ArgumentNullException("cellFormat");

            string numFormat = cellInfo.NumberFormat ?? formatCode;

            // Anything to apply?
            if (string.IsNullOrEmpty(numFormat)) return;

            // Look up existing...
            var item = stylesManager.numberFormats.Find(numFormat);
            UInt32Value id = item.Value;

            if (id == null)
            {
                // Create (and append?) a new numbering format in the stylesheet
                id = stylesheet.AddNumberingFormat(cellInfo.NumberFormat ?? formatCode);
                stylesManager.numberFormats.Add(numFormat, id);
            }

            cellFormat.NumberFormatId = id;
        }

        /// <summary>
        /// Updates the supplied <see cref="CellFormat"/> with a colour index which relates to the Fill Colour.
        /// The colour is created in the excel <see cref="Stylesheet"/> if it is not already there.
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="cellInfo"></param>
        /// <returns></returns>
        private static void UpdateBorder(ExcelStylesManager stylesManager, ExcelCellStyleInfo cellInfo, ref Stylesheet stylesheet, ref CellFormat cellFormat)
        {
            if (cellFormat == null) throw new ArgumentNullException("cellFormat");

            if (cellInfo.BorderInfo != null && cellInfo.BorderInfo.HasBorder)
            {
                var item = stylesManager.borders.Find(cellInfo.BorderInfo);
                UInt32Value id = item.Value;

                if (id == null)
                {
                    // Add and return the index of a fill colour
                    id = CreateNewBorder(stylesheet, cellInfo.BorderInfo);
                    stylesManager.borders.Add(cellInfo.BorderInfo, id);
                }
                cellFormat.BorderId = id;
            }
        }

        /// <summary>
        /// Creates and appends a new border to the supplied <see cref="Stylesheet"/> based on properties of the <see cref="ExcelCellBorderInfo"/>.<br/>
        /// Returns the index of the newly created border.
        /// </summary>
        /// <param name="stylesheet"></param>
        /// <param name="borderInfo"></param>
        /// <returns></returns>
        private static uint CreateNewBorder(Stylesheet stylesheet, ExcelCellBorderInfo borderInfo)
        {
            Border border = new Border();

            // LEFT
            if (borderInfo.HasLeftBorder)
            {
                border.LeftBorder = new LeftBorder()
                {
                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
                    {
                        Rgb = new HexBinaryValue(StyleTranslator.Translate(borderInfo.ColourLeft.Value))
                    },
                    Style = StyleTranslator.Translate(borderInfo.WidthLeft)
                };
            }

            // TOP
            if (borderInfo.HasTopBorder)
            {
                border.TopBorder = new TopBorder()
                {
                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
                    {
                        Rgb = new HexBinaryValue(StyleTranslator.Translate(borderInfo.ColourTop.Value))
                    },
                    Style = StyleTranslator.Translate(borderInfo.WidthTop)
                };
            }

            // RIGHT
            if (borderInfo.HasRightBorder)
            {
                border.RightBorder = new RightBorder()
                {
                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
                    {
                        Rgb = new HexBinaryValue(StyleTranslator.Translate(borderInfo.ColourRight.Value))
                    },
                    Style = StyleTranslator.Translate(borderInfo.WidthRight)
                };
            }

            // BOTTOM
            if (borderInfo.HasBottomBorder)
            {
                border.BottomBorder = new BottomBorder()
                {
                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color()
                    {
                        Rgb = new HexBinaryValue(StyleTranslator.Translate(borderInfo.ColourBottom.Value))
                    },
                    Style = StyleTranslator.Translate(borderInfo.WidthBottom)
                };
            }

            stylesheet.Borders.Append(border);
            uint borderId = (uint)new List<object>(stylesheet.Borders.Cast<object>()).IndexOf(border);
            return borderId;
        }

        /// <summary>
        /// Create an Allignment object which manages indentation, alignment and text rotation and applies it to the cellFormat.
        /// </summary>
        /// <param name="cellInfo"></param>
        /// <param name="stylesheet"></param>
        /// <param name="cellFormat"></param>
        private static void UpdateTextAlignmentAndRotation(ExcelCellStyleInfo cellInfo, ref CellFormat cellFormat)
        {
            if (cellFormat == null) throw new ArgumentNullException("cellFormat");

            if (cellInfo.AlignmentInfo != null)
            {
                cellFormat.Alignment = new Alignment()
                {
                    Indent = new UInt32Value((uint)cellInfo.AlignmentInfo.LeftMargin / 2),
                    Horizontal = new EnumValue<HorizontalAlignmentValues>(StyleTranslator.Translate(cellInfo.AlignmentInfo.TextAlignment)),
                    Vertical = new EnumValue<VerticalAlignmentValues>(StyleTranslator.Translate(cellInfo.AlignmentInfo.VerticalAlignment)),
                    WrapText = new BooleanValue(cellInfo.AlignmentInfo.TextWrapping != TextWrapping.NoWrap),
                    TextRotation = (uint)cellInfo.AlignmentInfo.TextRotationAngle,
                };
            }
        }

        #endregion Private Helpers

        internal StyleBase GetMapStyle(string key)
        {
            return this.GetMapStyle(key, null);
        }

        internal StyleBase GetMapStyle(string key, object dataContext)
        {
            if (string.IsNullOrEmpty(key))
            {
                return null;
            }

            //  Inflate and merge as required
            var mapStyle = this.ResourceStore.GetResourceByKey<StyleBase>(key);
            mapStyle.DataContext = dataContext;

            mapStyle = this.Merge(mapStyle);
            return mapStyle;

            // Use the dictionary of Current Map Styles to perform the translation
            // return this.currentMapStyles.FindByKey(key);
        }
    }
}

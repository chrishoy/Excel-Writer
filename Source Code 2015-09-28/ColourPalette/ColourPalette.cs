namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Windows.Media;

    /// <summary>
    /// Colour palette as specified in GAM Standards reference document, which can be found here:<br />
    /// http://msas-ldn29t/sites/MM/General/Development/Look-and-feel/GAM Visual Identity.pdf
    /// </summary>
    public static class ColourPalette
    {
        /// <summary>
        /// The chart brush list
        /// </summary>
        private static List<Brush> chartBrushList = new List<Brush>();
        /// <summary>
        /// The colour palettes
        /// </summary>
        private static Dictionary<ColourPaletteType, GamPalette> colourPalettes = new Dictionary<ColourPaletteType,GamPalette>();
        /// <summary>
        /// The internal brush list
        /// </summary>
        private static ReadOnlyCollection<Brush> internalBrushList;

        #region Constructor

        /// <summary>
        /// fred
        /// </summary>
        static ColourPalette()
        {
            var allColoursPalette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.AllPalette, allColoursPalette);

            // Loads with GAM Standard Colours
            LoadChartPalette();
            LoadBrandPalette();
            LoadProductPalette();
            LoadTechnicalPalette();
            LoadGeneralPalette();
            LoadGreyScalePalette();
            LoadTechnicalChartPalette();

            // Populates the chartBrushList (CHART COLOUR PALETTE ONLY for backward compatibility).
            PopulateBrushList(ref chartBrushList, colourPalettes[ColourPaletteType.ChartPalette].ColorList);
        }

        #endregion Constructor

        #region Private Helpers

        /// <summary>
        /// Populates the offical GAM technical palette (Includes GAM Greens).
        /// Assigns them to groups.
        /// </summary>
        private static void LoadTechnicalPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GamTechnicalPalette, palette);

            // 1 to 10
            AppendColour(palette, Color.FromRgb(0, 102, 102), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(94, 146, 216), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(141, 134, 130), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(190, 144, 69), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(113, 130, 70), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(188, 82, 155), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(208, 99, 81), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(188, 172, 72), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(121, 88, 148), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(0, 167, 139), ColourGroup.GamTPAqua);

            // 11 to 20
            AppendColour(palette, Color.FromRgb(102, 163, 163), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(158, 190, 232), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(187, 182, 180), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(216, 188, 143), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(170, 180, 144), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(215, 151, 195), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(227, 161, 151), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(215, 205, 145), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(175, 155, 191), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(102, 202, 185), ColourGroup.GamTPAqua);

            // 21 to 30 - a bit wishy washy
            AppendColour(palette, Color.FromRgb(178, 209, 209), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(207, 222, 243), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(221, 219, 217), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(235, 222, 199), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(212, 217, 199), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(235, 203, 225), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(241, 208, 203), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(235, 230, 200), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(215, 205, 223), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(178, 229, 220), ColourGroup.GamTPAqua);
        }

        /// <summary>
        /// Populates the offical technical palette which should really be used for charts (but some colours are 'wishy-washy').
        /// Assigns them to groups.
        /// </summary>
        private static void LoadTechnicalChartPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GamTechnicalChartPalette, palette);

            // 60% tinted colours first
            AppendColour(palette, Color.FromRgb(158, 190, 232), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(187, 182, 180), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(216, 188, 143), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(170, 180, 144), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(215, 151, 195), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(227, 161, 151), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(215, 205, 145), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(175, 155, 191), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(102, 202, 185), ColourGroup.GamTPAqua);

            // 30% tinted colours second
            AppendColour(palette, Color.FromRgb(207, 222, 243), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(221, 219, 217), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(235, 222, 199), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(212, 217, 199), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(235, 203, 225), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(241, 208, 203), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(235, 230, 200), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(215, 205, 223), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(178, 229, 220), ColourGroup.GamTPAqua);

            // 100% solid colours third
            AppendColour(palette, Color.FromRgb(94, 146, 216), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(141, 134, 130), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(190, 144, 69), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(113, 130, 70), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(188, 82, 155), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(208, 99, 81), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(188, 172, 72), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(121, 88, 148), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(0, 167, 139), ColourGroup.GamTPAqua);
        
        }

        /// <summary>
        /// Populates the (unofficial) Chart colour palette.<br />
        /// This is based on the 'Technical Palette' but removes the 'wishy-washy' colours.
        /// Assigns them to groups.
        /// </summary>
        private static void LoadChartPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.ChartPalette, palette);
            
            // 1 to 10
            AppendColour(palette, Color.FromRgb(0, 102, 102), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(94, 146, 216), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(141, 134, 130), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(190, 144, 69), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(113, 130, 70), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(188, 82, 155), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(208, 99, 81), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(188, 172, 72), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(121, 88, 148), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(0, 167, 139), ColourGroup.GamTPAqua);

            // 11 to 20
            AppendColour(palette, Color.FromRgb(102, 163, 163), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(158, 190, 232), ColourGroup.GamTPBlue);
            AppendColour(palette, Color.FromRgb(187, 182, 180), ColourGroup.GamTPGrey);
            AppendColour(palette, Color.FromRgb(216, 188, 143), ColourGroup.GamTPMustard);
            AppendColour(palette, Color.FromRgb(170, 180, 144), ColourGroup.GamTPOlive);
            AppendColour(palette, Color.FromRgb(215, 151, 195), ColourGroup.GamTPRubine);
            AppendColour(palette, Color.FromRgb(227, 161, 151), ColourGroup.GamTPOrange);
            AppendColour(palette, Color.FromRgb(215, 205, 145), ColourGroup.GamTPGold);
            AppendColour(palette, Color.FromRgb(175, 155, 191), ColourGroup.GamTPPurple);
            AppendColour(palette, Color.FromRgb(102, 202, 185), ColourGroup.GamTPAqua);

            // 21 to 30 - a bit wishy washy removed for now
            //AppendColour(palette, 178, 209, 209, GamColourGroup.GamGreen);
            //AppendColour(palette, 207, 222, 243, GamColourGroup.GamTPBlue);
            //AppendColour(palette, 221, 219, 217, GamColourGroup.GamTPGrey);
            //AppendColour(palette, 235, 222, 199, GamColourGroup.GamTPMustard);
            //AppendColour(palette, 212, 217, 199, GamColourGroup.GamTPOlive);
            //AppendColour(palette, 235, 203, 225, GamColourGroup.GamTPRubine);
            //AppendColour(palette, 241, 208, 203, GamColourGroup.GamTPOrange);
            //AppendColour(palette, 235, 230, 200, GamColourGroup.GamTPGold);
            //AppendColour(palette, 215, 205, 223, GamColourGroup.GamTPPurple);
            //AppendColour(palette, 178, 229, 220, GamColourGroup.GamTPAqua);
        }

        /// <summary>
        /// Populates the GAM Brand colour palette (Gam Greens)
        /// Assigns them to groups.
        /// </summary>
        private static void LoadBrandPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GamBrandPalette, palette);

            AppendColour(palette, Color.FromRgb(0, 102, 102), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(102, 163, 163), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(178, 209, 209), ColourGroup.GamGreen);
            AppendColour(palette, Color.FromRgb(242, 247, 247), ColourGroup.GamGreen);
        }


        /// <summary>
        /// Populates the GAM Product colour palette (Products)
        /// Assigns them to groups.
        /// </summary>
        private static void LoadProductPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GamProductPalette, palette);

            AppendColour(palette, Color.FromRgb(54, 183, 229), ColourGroup.GamPPLightBlue);
            AppendColour(palette, Color.FromRgb(151, 197, 36), ColourGroup.GamPPLime);
            AppendColour(palette, Color.FromRgb(255, 103, 109), ColourGroup.GamPPRed);
            AppendColour(palette, Color.FromRgb(245, 140, 66), ColourGroup.GamPPTangerine);
            AppendColour(palette, Color.FromRgb(171, 128, 203), ColourGroup.GamPPViolet);
            AppendColour(palette, Color.FromRgb(110, 119, 200), ColourGroup.GamPPPlum);
        }

        /// <summary>
        /// Populates the General Palette. GreyScale, Other and GamGreenPlus (all non-GAM Standard)
        /// Assigns them to groups.
        /// </summary>
        private static void LoadGeneralPalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GeneralPalette, palette);
            
            AppendColour(palette, Color.FromRgb(255, 0, 0), ColourGroup.Others);           // Red
            AppendColour(palette, Color.FromRgb(0, 255, 0), ColourGroup.Others);           // Green
            AppendColour(palette, Color.FromRgb(0, 0, 255), ColourGroup.Others);           // Blue

            AppendColour(palette, Color.FromRgb(0, 0, 0), ColourGroup.GreyScale);          // Black
            AppendColour(palette, Color.FromRgb(51, 51, 51), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(102, 102, 102), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(140, 140, 140), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(178, 178, 178), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(216, 216, 216), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(255, 255, 255), ColourGroup.GreyScale);    // White

            AppendColour(palette, Color.FromRgb(0, 102, 102), ColourGroup.GamGreenPlus);   // GAM Green 100%
            AppendColour(palette, Color.FromRgb(51, 132, 132), ColourGroup.GamGreenPlus);  
            AppendColour(palette, Color.FromRgb(102, 163, 163), ColourGroup.GamGreenPlus); // GAM Green 60%
            AppendColour(palette, Color.FromRgb(140, 186, 186), ColourGroup.GamGreenPlus);
            AppendColour(palette, Color.FromRgb(178, 209, 209), ColourGroup.GamGreenPlus); // GAM Green 30%
            AppendColour(palette, Color.FromRgb(216, 232, 232), ColourGroup.GamGreenPlus);
            AppendColour(palette, Color.FromRgb(242, 247, 247), ColourGroup.GamGreenPlus); // GAM Green 5%

            AppendColour(palette, Color.FromRgb(255, 255, 255), ColourGroup.GamGreenPlus); // White
        }

        /// <summary>
        /// Populates a GreyScale - non-GAM palette.
        /// Assigns them to groups.
        /// </summary>
        private static void LoadGreyScalePalette()
        {
            var palette = new GamPalette();
            colourPalettes.Add(ColourPaletteType.GreyScalePalette, palette);

            AppendColour(palette, Color.FromRgb(0, 0, 0), ColourGroup.GreyScale);          // Black
            AppendColour(palette, Color.FromRgb(51, 51, 51), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(102, 102, 102), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(140, 140, 140), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(178, 178, 178), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(216, 216, 216), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(247, 247, 247), ColourGroup.GreyScale);
            AppendColour(palette, Color.FromRgb(255, 255, 255), ColourGroup.GreyScale);    // White
        }

        /// <summary>
        /// Populates a supplied brush list with colours in the supplied colour list.
        /// </summary>
        /// <param name="brushList">List to be populated</param>
        /// <param name="colourList">Source colour list</param>
        /// <exception cref="ArgumentNullException">brushList</exception>
        private static void PopulateBrushList(ref List<Brush> brushList, List<Color> colourList)
        {
            if (brushList == null) throw new ArgumentNullException("brushList");

            brushList.Clear();
            foreach (Color colour in colourList)
            {
                brushList.Add(new SolidColorBrush(colour));
            }
        }

        /// <summary>
        /// Returns a brush list
        /// </summary>
        /// <value>
        /// The brush list.
        /// </value>
        public static ReadOnlyCollection<Brush> BrushList
        {
            get
            {
                if (internalBrushList == null)
                {
                    internalBrushList = new ReadOnlyCollection<Brush>(chartBrushList);
                }
                return internalBrushList;
            }
        }

        /// <summary>
        /// Gets a new, fully pupulated, brush list containing the GAM Chart Palette (note this is not the Technical Chart Palette).
        /// </summary>
        /// <returns>
        /// A brush list.
        /// </returns>
        public static IEnumerable<Brush> GetNewBrushList()
        {
            var newBrushList = new List<Brush>();
            PopulateBrushList(ref newBrushList, colourPalettes[ColourPaletteType.ChartPalette].ColorList);
            return newBrushList;            
        }

        /// <summary>
        /// Creates and returns an existing instance of a brush IN THE CHART COLOUR PALETTE based on a zero-based index colour.<br />
        /// If the index exceeds the number of elements in the palette, then a remainder index is used instead.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        [Obsolete("This causes cross-threading issues. Use GetNewBrush instead.")]
        public static Brush GetBrush(int index)
        {
            int remainder;
            Math.DivRem(index, chartBrushList.Count, out remainder);

            return chartBrushList[remainder];
        }

        /// <summary>
        /// Creates and returns a new instance of a brush IN THE CHART COLOUR PALETTE based on a zero-based index colour.<br />
        /// The brush list represents the chart palette.<br />
        /// If the index exceeds the number of elements in the palette, then a remainder index is used instead.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        [Obsolete("Used GetNewBrush, specifying a PaletteType of (eg.) ChartPalette instead")]
        public static Brush GetNewBrush(int index)
        {
            int remainder;
            List<Color> cl = colourPalettes[ColourPaletteType.ChartPalette].ColorList;
            Math.DivRem(index, cl.Count, out remainder);
            return new SolidColorBrush(cl[remainder]);
        }

        /// <summary>
        /// Creates and returns a new instance of a brush based on a zero-based index within a palette.<br />
        /// If the index exceeds the number of elements in the palette, then a remainder index is used instead.
        /// </summary>
        /// <param name="palette">The <see cref="ColourPaletteType" /></param>
        /// <param name="index">The index used to look up a colour</param>
        /// <returns>
        /// A new instance of a <see cref="SolidColorBrush" />
        /// </returns>
        public static Brush GetNewBrush(ColourPaletteType palette, int index)
        {
            return new SolidColorBrush(GetColour(palette, index));
        }

        /// <summary>
        /// Creates and returns a new instance of a brush based on a zero-based index within a colour group, within a palette.<br />
        /// If the index exceeds the number of elements in the colour group, then a remainder index is used instead.
        /// </summary>
        /// <param name="palette">The <see cref="ColourPaletteType" /></param>
        /// <param name="colourGroup">The <see cref="ColourGroup" /></param>
        /// <param name="index">The index used to look up a colour</param>
        /// <returns>
        /// A new instance of a <see cref="SolidColorBrush" />
        /// </returns>
        public static Brush GetNewBrush(ColourPaletteType palette, ColourGroup colourGroup, int index)
        {
            return new SolidColorBrush(GetColour(palette, colourGroup, index));
        }

        /// <summary>
        /// Gets a colour based on a zero-based index within a colour group, within a palette.<br />
        /// If the index exceeds the number of elements in the colour group, then a remainder index is used instead.
        /// </summary>
        /// <param name="palette">The <see cref="ColourPaletteType" /></param>
        /// <param name="colourGroup">The <see cref="ColourGroup" /></param>
        /// <param name="index">The index used to look up a colour</param>
        /// <returns>
        /// A <see cref="Color" />
        /// </returns>
        public static Color GetColour(ColourPaletteType palette, ColourGroup colourGroup, int index)
        {
            int remainder;
            var gp = colourPalettes[palette].ColourGroups;

            List<Color> colourGroupColours = colourPalettes[palette].ColourGroups[colourGroup];
            Math.DivRem(index, colourGroupColours.Count, out remainder);

            return colourGroupColours[remainder];
        }

        /// <summary>
        /// Gets a colour based on a zero-based index within a palette.<br />
        /// If the index exceeds the number of elements in the palette, then a remainder index is used instead.
        /// </summary>
        /// <param name="palette">The <see cref="ColourPaletteType" /></param>
        /// <param name="index">The index used to look up a colour</param>
        /// <returns>
        /// A <see cref="Color" />
        /// </returns>
        public static Color GetColour(ColourPaletteType palette, int index)
        {
            int remainder;
            List<Color> colourList = colourPalettes[palette].ColorList;
            Math.DivRem(index, colourList.Count, out remainder);
            return colourList[remainder];
        }

        /// <summary>
        /// Returns the number of colours in the specified palette, so that we can (for example)
        /// request them in reverse order.
        /// </summary>
        /// <param name="palette">The palette.</param>
        /// <returns></returns>
        public static int CountColours(ColourPaletteType palette)
        {
            return colourPalettes[palette].ColorList.Count;
        }

        #endregion Private Helpers

        /// <summary>
        /// Append the colour to the appropriate group and to the main list of colours in the supplied <see cref="gamPalette" />
        /// </summary>
        /// <param name="gamPalette">The gam palette.</param>
        /// <param name="colour">The colour.</param>
        /// <param name="gamColourGroup">The gam colour group.</param>
        internal static void AppendColour(GamPalette gamPalette, Color colour, ColourGroup gamColourGroup)
        {
            // Append to supplied palette.
            AppendToPalette(gamPalette, colour, gamColourGroup);

            // Add to 'All Colours' palette.
            AppendToPalette(colourPalettes[ColourPaletteType.AllPalette], colour, gamColourGroup);

        }

        /// <summary>
        /// Appends to palette.
        /// </summary>
        /// <param name="gamPalette">The gam palette.</param>
        /// <param name="colour">The colour.</param>
        /// <param name="gamColourGroup">The gam colour group.</param>
        private static void AppendToPalette(GamPalette gamPalette, Color colour, ColourGroup gamColourGroup)
        {
            // Check if colour group exists, if not then create
            if (!gamPalette.ColourGroups.ContainsKey(gamColourGroup))
            {
                gamPalette.ColourGroups.Add(gamColourGroup, new List<Color>());
            }

            gamPalette.ColourGroups[gamColourGroup].Add(colour);

            // Add to full list if not already there.
            if (!gamPalette.ColorList.Exists(c => c.Equals(colour)))
            {
                gamPalette.ColorList.Add(colour);
            }
            else
            {

            }
        }
    }
}

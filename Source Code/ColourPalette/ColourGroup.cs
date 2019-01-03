namespace ExcelWriter
{
    /// <summary>
    /// Determines the colour group in which a set of colours belong.
    /// </summary>
    public enum ColourGroup
    {
        /// <summary>
        /// GAM Green 
        /// </summary>
        GamGreen,

        /// <summary>
        /// GAM Technical Palette Blue
        /// </summary>
        GamTPBlue,

        /// <summary>
        /// GAM Technical Palette Grey
        /// </summary>
        GamTPGrey,

        /// <summary>
        /// GAM Technical Palette Mustard
        /// </summary>
        GamTPMustard,

        /// <summary>
        /// GAM Technical Palette Olive
        /// </summary>
        GamTPOlive,

        /// <summary>
        /// GAM Technical Palette Rubine (Magenta)
        /// </summary>
        GamTPRubine,

        /// <summary>
        /// GAM Technical Palette Orange
        /// </summary>
        GamTPOrange,

        /// <summary>
        /// GAM Technical Palette Gold
        /// </summary>
        GamTPGold,

        /// <summary>
        /// GAM Technical Palette Purple
        /// </summary>
        GamTPPurple,

        /// <summary>
        /// GAM Technical Palette Aqua
        /// </summary>
        GamTPAqua,

        /// <summary>
        /// GAM Product Palette Light Blue (used for 'Single Manager Long Only Funds')
        /// </summary>
        GamPPLightBlue,

        /// <summary>
        /// GAM Product Palette Light Blue (used for 'Single Manager Hedge Funds')
        /// </summary>
        GamPPLime,

        /// <summary>
        /// GAM Product Palette Light Blue (used for 'Fund of Hedge Funds')
        /// </summary>
        GamPPRed,

        /// <summary>
        /// GAM Product Palette Tangerine (used for 'Composite Absolute Return Funds')
        /// </summary>
        GamPPTangerine,

        /// <summary>
        /// GAM Product Palette Violet (used for 'GAM Structured Investments')
        /// </summary>
        GamPPViolet,

        /// <summary>
        /// GAM Product Palette Plum (used for 'Single Manager Long/Short Funds')
        /// </summary>
        GamPPPlum,

        /// <summary>
        /// Non-standard colours - blacks/greys 
        /// </summary>
        GreyScale,

        /// <summary>
        /// Non-standard GAM Greens (there is an extra scale of greens between the official GAM Greens to allow for extra banding in reports).
        /// </summary>
        GamGreenPlus,

        /// <summary>
        /// Other colours, such as black, white, green, red, blue.
        /// </summary>
        Others
    }
}

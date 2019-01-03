namespace ExcelWriter
{
    /// <summary>
    /// Different types of colour palette used within GAM.
    /// </summary>
    public enum ColourPaletteType
    {
        /// <summary>
        /// Unofficial colour palette which is based on the GAM TechnicalPalette - Plus GAM Greens, but excludes the lightest shades.
        /// </summary>
        ChartPalette,

        /// <summary>
        /// Full official GAM technical colour palette. Includes GAM Greens.
        /// </summary>
        GamTechnicalPalette,

        /// <summary>
        /// Official GAM technical colour palette when used for creating charts.
        /// </summary>
        GamTechnicalChartPalette,

        /// <summary>
        /// GAM Greens, official Brand Colour palette.
        /// </summary>
        GamBrandPalette,

        /// <summary>
        /// GAM Products colour palette. Includes colours for 'Single Manager Long Only Funds', 'Single Manager Hedge Funds', 'Funds of Hedge Funds' etc...'
        /// </summary>
        GamProductPalette,

        /// <summary>
        /// Other colours which happen to be missing from the GAM Palettes, such as White, Black, Red, Blue, Green.<br/>
        /// Also includes 'Grey-Scale' and 'Gam-Green' Colour Groups.
        /// </summary>
        GeneralPalette,

        /// <summary>
        /// All of the colours, including GAM and non-GAM
        /// </summary>
        AllPalette,

        /// <summary>
        /// All of the grey-scale colours, these are non-GAM
        /// </summary>
        GreyScalePalette,
    }
}

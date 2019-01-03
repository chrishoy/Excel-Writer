namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Windows.Media;

    /// <summary>
    /// Represents a palette of GAM colours
    /// </summary>
    internal class GamPalette
    {
        #region Local Fields

        private List<Color> colorList = new List<Color>();
        private Dictionary<ColourGroup, List<Color>> colourGroups;

        #endregion Local Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="GamPalette" /> class.
        /// </summary>
        public GamPalette()
        {
            this.colorList = new List<Color>();
            this.colourGroups = new Dictionary<ColourGroup, List<Color>>();
        }

        #endregion Construction

        /// <summary>
        /// Gets a full list of colours within this palette
        /// </summary>
        public List<Color> ColorList
        {
            get { return this.colorList; }
        }

        /// <summary>
        /// Gets a dictionary of colours within this palette, grouped by colour group.
        /// </summary>
        public Dictionary<ColourGroup, List<Color>> ColourGroups
        {
            get { return this.colourGroups; }
        }
    }
}

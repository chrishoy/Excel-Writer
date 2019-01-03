namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a row co-ordinate within a <see cref="ExcelMapCoOrdinateContainer"/>.<br/>
    /// Used for managing the content of <see cref="ExcelMapCoOrdinateContainer"/> when writing to Excel.
    /// </summary>
    internal class ExcelMapCoOrdinateRow
    {
        #region Private Fields

        private uint worksheetRowIndex;
        private double? height;
        private bool isHidden;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Gets or sets a height for the row. Null indicates not specified.
        /// </summary>
        public double? Height
        {
            get { return this.height; }
            set { this.height = value; }
        }

        /// <summary>
        /// Gets or sets the index of the row in the final worksheet.
        /// </summary>
        public uint WorksheetRowIndex
        {
            get { return this.worksheetRowIndex; }
            set { this.worksheetRowIndex = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the column is hidden
        /// </summary>
        public bool IsHidden
        {
            get { return this.isHidden; }
            set { this.isHidden = value; }
        }

        #endregion Public Properties

        #region Overrides

        /// <summary>
        /// Returns a string representation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            return string.Format("ExcelMapCoOrdinateRow:WorksheetRowIndex={0},Height={1},Hidden={2}",
                                 this.worksheetRowIndex,
                                 this.height,
                                 this.isHidden);
        }

        #endregion Overrides
    }
}

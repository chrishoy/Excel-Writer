namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a single cell co-ordinate within a <see cref="ExcelMapCoOrdinateContainer"/>.<br/>
    /// Used for managing <see cref="StackPanel"/> to Excel cell writing.
    /// </summary>
    internal class ExcelMapCoOrdinateCell : ExcelMapCoOrdinatePlaceholder
    {
        #region Private Fields

        private object cellValue;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapCoOrdinateCell" /> class.
        /// </summary>
        public ExcelMapCoOrdinateCell()
            : base()
        {
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the value to be written into the cell
        /// </summary>
        public object CurrentValue
        {
            get { return this.cellValue; }
            set { this.cellValue = value; }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Returns a string represenation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            //string totRowCount = this.RowsCounted ? string.Format("{0}", this.TotalRowCount) : "?";
            //string startRowIdx = this.RowsCounted ? string.Format("{0}", this.StartRowIndex) : "?";

            return string.Format
                (
                    "ExcelMapCoOrdinateCell[Id={0}]:WorksheetRowCol=R{1}C{2}",
                    this.Id,
                    this.ExcelRowStart,
                    this.ExcelColumnStart
                );
        }

        #endregion Public Methods

        #region Internal Methods

        #endregion Internal Methods
    }
}

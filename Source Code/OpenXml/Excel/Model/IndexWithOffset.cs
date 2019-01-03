namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a Row/Column in Excel with associated offsets in emus.
    /// </summary>
    public class RowColumnIndexWithOffset
    {
        #region Construction

        /// <summary>
        /// Default constructor
        /// </summary>
        public RowColumnIndexWithOffset()
        {
            this.Row = new IndexOffset();
            this.Column = new IndexOffset();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public RowColumnIndexWithOffset(uint rowIndex, uint columnIndex)
        {
            this.Row = new IndexOffset { Index = rowIndex };
            this.Column = new IndexOffset { Index = rowIndex };
        }

        #endregion Construction

        #region Public Porperties

        public IndexOffset Row { get; set; }
        public IndexOffset Column { get; set; }

        #endregion Public Properties

        #region Public Methods

        public override string ToString()
        {
            return string.Format("Row={0},Column={1}", this.Row, this.Column);
        }

        #endregion Public Methods

    }
}

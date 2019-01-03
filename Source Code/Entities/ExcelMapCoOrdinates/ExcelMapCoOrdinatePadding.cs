namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a padding co-ordinate within a <see cref="ExcelMapCoOrdinateContainer"/>.<br/>
    /// Used for creating padding cells when writing to Excel.
    /// </summary>
    internal class ExcelMapCoOrdinatePadding : ExcelMapCoOrdinate
    {
        #region Private Fields

        private object cellValue;

        private uint columnSpan;
        private uint rowSpan;
        //private uint startRowIndex;
        //private uint endRowIndex;
        //private uint totalRowCount;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapCoOrdinatePadding" /> class.
        /// </summary>
        public ExcelMapCoOrdinatePadding()
            : base()
        {
            this.columnSpan = 0;
            this.rowSpan = 0;
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

        /// <summary>
        /// Gets or sets the number of columns within it's container that the padding will span.<br/>
        /// Default value is 0.
        /// </summary>
        public uint ColumnSpan
        {
            get { return this.columnSpan; }
            set { this.columnSpan = value; }
        }

        /// <summary>
        /// Gets or sets the number of rows within it's container that the padding will span.<br/>
        /// Default value is 0.
        /// </summary>
        public uint RowSpan
        {
            get { return this.rowSpan; }
            set { this.rowSpan = value; }
        }

        /// <summary>
        /// Gets the number of rows this entity represents.
        /// </summary>
        public override uint MapRowCount
        {
            get { return 1; }
        }

        /// <summary>
        /// Gets the number of columns this entity represents
        /// </summary>
        public override uint MapColumnCount
        {
            get { return 1; }
        }

        ///// <summary>
        ///// Gets the number of rows that this entity represents in an Excel worksheet
        ///// </summary>
        //public override uint TotalRowCount
        //{
        //    get
        //    {
        //        if (!this.RowsCounted)
        //        {
        //            throw new InvalidOperationException("TotalRowCount has not yet been calculated. Call CalculateTotalRowCount first.");
        //        }

        //        return this.totalRowCount;
        //    }
        //}

        ///// <summary>
        ///// Gets the index of Excel row where this entity starts in the excel worksheet
        ///// </summary>
        //public override uint StartRowIndex
        //{
        //    get { return this.startRowIndex; }
        //}

        ///// <summary>
        ///// Index of excel row where this entity ends in the excel worksheet
        ///// </summary>
        //public override uint EndRowIndex
        //{
        //    get { return this.endRowIndex; }
        //}

        // <summary>
        // Gets the 1-based index of excel row where this entity starts in the excel worksheet
        // </summary>
        public override int ExcelRowStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel row where this entity ends in the excel worksheet
        /// </summary>
        public override int ExcelRowEnd { get; set; }

        // <summary>
        // Gets the 1-based index of excel column where this entity starts in the excel worksheet
        // </summary>
        public override int ExcelColumnStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity ends in the excel worksheet
        /// </summary>
        public override int ExcelColumnEnd { get; set; }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Returns a string represenation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            //string totRowCount = this.RowsCounted ? string.Format("{0}", this.TotalRowCount) : "?";
            //string startRowIdx = this.RowsCounted ? string.Format("{0}", this.startRowIndex) : "?";

            return string.Format
                (
                    "ExcelMapCoOrdinatePadding[Id={0}]:WorksheetRowCol=R{1}C{2}",
                    this.Id,
                    this.ExcelRowStart,
                    this.ExcelColumnStart
                );
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Builds a model of the columns within this entity.
        /// </summary>
        internal override RowOrColumnsModel BuildColumnsModel()
        {
            return new RowOrColumnsModel(false);
        }

        /// <summary>
        /// Builds a model of the rows within this entity.
        /// </summary>
        internal override RowOrColumnsModel BuildRowsModel()
        {
            return new RowOrColumnsModel(true);
        }

        ///// <summary>
        ///// Calculates, stores and returns the total rows count that this entity represents in an excel worksheet. (i.e. 1)
        ///// </summary>
        ///// <param name="startRowIndex">The Excel row where this entity starts</param>
        ///// <returns>The total row count that this entity represents in an excel worksheet.</returns>
        //internal override uint CountRows(uint startRowIndex)
        //{
        //    this.startRowIndex = startRowIndex;
        //    this.totalRowCount = this.rowSpan; // == 0 ? 1 : this.rowSpan; // 1
        //    this.endRowIndex = this.startRowIndex + this.totalRowCount - 1; // CHANGE

        //    this.RowsCounted = true;
        //    return this.totalRowCount;
        //}

        #endregion Internal Methods
    }
}

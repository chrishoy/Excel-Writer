namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a placeholder co-ordinate within a <see cref="ExcelMapCoOrdinateContainer"/>.<br/>
    /// Used for placing charts in Excel cells when writing to Excel.
    /// </summary>
    internal class ExcelMapCoOrdinatePlaceholder : ExcelMapCoOrdinate
    {
        #region Private Fields

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapCoOrdinatePlaceholder" /> class.
        /// </summary>
        public ExcelMapCoOrdinatePlaceholder()
            : base()
        {
            this.ColumnSpan = 1;
            this.RowSpan = 1;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the number of columns within it's container that the cell will span.<br/>
        /// Default value is 1
        /// </summary>
        public uint ColumnSpan { get; set; }

        /// <summary>
        /// Gets or sets the number of rows within it's container that the cell will span.<br/>
        /// Default value is 1
        /// </summary>
        public uint RowSpan { get; set; }

        /// <summary>
        /// Gets or sets a cell which this cell should be merged into.
        /// </summary>
        public ExcelMapCoOrdinatePlaceholder MergeWith { get; set; }

        /// <summary>
        /// Gets the number of rows this entity represents
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

        /// <summary>
        /// Gets the 1-based index of excel row where this entity starts in the excel worksheet
        /// </summary>
        public override int ExcelRowStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel row where this entity ends in the excel worksheet
        /// </summary>
        public override int ExcelRowEnd { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity starts in the excel worksheet
        /// </summary>
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
            return string.Format
                (
                    "ExcelMapCoOrdinatePlaceholder:WorksheetRowCol=R{0}C{1}",
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
            var columnsModel = new RowOrColumnsModel(false);
            columnsModel.Add(this);

            if (this.ColumnSpan > 1)
            {
                for (int count = 2; count <= this.ColumnSpan; count++)
                {
                    columnsModel.Add(this);
                }
            }

            return columnsModel;
        }

        /// <summary>
        /// Builds a model of the rows within this entity.
        /// </summary>
        internal override RowOrColumnsModel BuildRowsModel()
        {
            var rowsModel = new RowOrColumnsModel(true);
            rowsModel.Add(this);

            if (this.RowSpan > 1)
            {
                for (int count = 2; count <= this.RowSpan; count++)
                {
                    rowsModel.Add(this);
                }
            }

            return rowsModel;
        }

        #endregion Internal Methods
    }
}

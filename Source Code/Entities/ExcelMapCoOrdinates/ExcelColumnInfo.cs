using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about a column that is to be written into an excel worksheet
    /// </summary>
    internal class ExcelColumnInfo
    {
        #region Private Fields

        private bool hasColumnInfo;
        private uint worksheetColumnIndex;
        private bool hasHiddenColumn;
        private double? maxColumnWidth;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="worksheetColumnIndex"></param>
        public ExcelColumnInfo(uint worksheetColumnIndex)
        {
            this.worksheetColumnIndex = worksheetColumnIndex;
        }

        ///// <summary>
        ///// Ctor. Creates and updates with supplied column information
        ///// </summary>
        ///// <param name="column"></param>
        //public ExcelColumnInfo(ExcelMapCoOrdinateColumn column)
        //{
        //    if (column == null) throw new ArgumentNullException("column");
        //    this.worksheetColumnIndex = column.WorksheetColumnIndex;
        //    this.Update(column);
        //}

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// The worksheet column that this information relates to.
        /// </summary>
        public uint WorksheetColumnIndex
        {
            get { return this.worksheetColumnIndex; }
        }

        /// <summary>
        /// Returns true if there is any column information used to format column in an excel worksheet
        /// </summary>
        public bool HasColumnInfo
        {
            get { return this.hasColumnInfo; }
            internal set { this.hasColumnInfo = value; }
        }

        /// <summary>
        /// Maximum width specified for a column
        /// </summary>
        public double? MaxColumnWidth
        {
            get { return this.maxColumnWidth; }
            internal set { this.maxColumnWidth = value; }
        }

        /// <summary>
        /// Are any columns marked as hidden
        /// </summary>
        public bool HasHiddenColumn
        {
            get { return this.hasHiddenColumn; }
            internal set { this.hasHiddenColumn = value; }
        }

        #endregion Public Properties

        #region Methods

        /// <summary>
        /// Compares and updates the maximum column width encountered.
        /// </summary>
        /// <param name="columnWidth"></param>
        private void UpdateMaxColumnWidth(double? columnWidth)
        {
            // Determine the maximum encountered column width
            if (columnWidth.HasValue && this.maxColumnWidth.GetValueOrDefault() < columnWidth.Value)
            {
                this.maxColumnWidth = columnWidth;
                this.hasColumnInfo = true;
            }
        }

        /// <summary>
        /// Compares and updates the HasHiddenColumn property.
        /// </summary>
        /// <param name="columnHidden"></param>
        private void UpdateHasHiddenColumn(bool columnHidden)
        {
            if (!this.hasHiddenColumn)
            {
                this.hasHiddenColumn = columnHidden;
                this.hasColumnInfo = true;
            }
        }

        #endregion Methods

    }
}

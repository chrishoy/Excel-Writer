using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about a row that is to be written into an excel worksheet
    /// </summary>
    internal class ExcelRowInfo
    {
        #region Private Fields

        private bool hasRowInfo;
        private uint worksheetRowIndex;
        private bool hasHiddenRow;
        private double? maxRowHeight;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="worksheetRowIndex"></param>
        public ExcelRowInfo(uint worksheetRowIndex)
        {
            this.worksheetRowIndex = worksheetRowIndex;
        }

        ///// <summary>
        ///// Ctor. Creates and updates with supplied Row information
        ///// </summary>
        ///// <param name="Row"></param>
        //public ExcelRowInfo(ExcelMapCoOrdinateRow row)
        //{
        //    if (row == null) throw new ArgumentNullException("row");
        //    this.worksheetRowIndex = row.WorksheetRowIndex;
        //    this.Update(row);
        //}

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// The worksheet Row that this information relates to.
        /// </summary>
        public uint WorksheetRowIndex
        {
            get { return this.worksheetRowIndex; }
        }

        /// <summary>
        /// Returns true if there is any Row information used to format Row in an excel worksheet
        /// </summary>
        public bool HasRowInfo
        {
            get { return this.hasRowInfo; }
            internal set { this.hasRowInfo = value; }
        }

        /// <summary>
        /// Gets or sets the Maximum Height specified for a Row
        /// </summary>
        public double? MaxRowHeight
        {
            get { return this.maxRowHeight; }
            internal set { this.maxRowHeight = value; }
        }

        /// <summary>
        /// Gets or set whether there are any Rows marked as hidden
        /// </summary>
        public bool HasHiddenRow
        {
            get { return this.hasHiddenRow; }
            internal set { this.hasHiddenRow = value; }
        }

        #endregion Public Properties
    }
}

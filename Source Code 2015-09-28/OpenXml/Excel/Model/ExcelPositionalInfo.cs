namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents row/column 'from' and 'to' indexes, plus from/to X and Y offsets as defined for OpenXml part placement.
    /// </summary>
    public class ExcelPositionalInfo
    {
        #region Construction

        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelPositionalInfo()
        {
            this.From = new RowColumnIndexWithOffset();
            this.To = new RowColumnIndexWithOffset();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fromRowIndex">The worksheet row where the picture will start</param>
        /// <param name="fromColumnIndex">The worksheet column where the picture will start</param>
        /// <param name="toRowIndex">The worksheet row when the picture will end</param>
        /// <param name="toColumnIndex">The worksheet column where the picture will end</param>
        public ExcelPositionalInfo(uint fromRowIndex, uint fromColumnIndex, uint toRowIndex, uint toColumnIndex)
        {
            this.From = new RowColumnIndexWithOffset(fromRowIndex, fromColumnIndex);
            this.To = new RowColumnIndexWithOffset(toRowIndex, toColumnIndex);
        }

        #endregion Construction

        #region Public Porperties

        public RowColumnIndexWithOffset From { get; set; }
        public RowColumnIndexWithOffset To { get; set; }

        #endregion Public Properties
    }
}

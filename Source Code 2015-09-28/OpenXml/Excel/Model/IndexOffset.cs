namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a Row or Column in Excel with associated offsets in emus.
    /// </summary>
    public class IndexOffset
    {
        #region Construction

        /// <summary>
        /// Default constructor
        /// </summary>
        public IndexOffset()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="index"></param>
        /// <param name="offsetEmus"></param>
        public IndexOffset(uint index, long offsetEmus)
        {
            this.Index = index;
            this.OffsetInEmus = offsetEmus;
        }

        #endregion Construction

        #region Public Porperties

        public uint Index { get; set; }
        public long OffsetInEmus { get; set; }

        public double OffsetCentimeters
        {
            get { return this.OffsetInEmus / Constants.EmusPerCentimeter; }
        }

        #endregion Public Properties

        #region Public Methods

        public override string ToString()
        {
            return string.Format("IndexOffset[Index={0},Offset={1:#,##0.00} emus]", this.Index, this.OffsetInEmus);
        }

        #endregion Public Methods

    }
}

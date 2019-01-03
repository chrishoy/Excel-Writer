namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents one or more <see cref="RangeReference"/>s.
    /// </summary>
    public class CompositeRangeReference
    {
        #region Private Fields

        private List<RangeReference> rangeReferences;
        private string sheetName;
        private uint minRowIndex;
        private uint maxRowIndex;
        private uint minColumnIndex;
        private uint maxColumnIndex;
        private string reference;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="CompositeRangeReference" /> class.<br/>
        /// This allows us to add range references together.
        /// </summary>
        public CompositeRangeReference()
        {
            this.rangeReferences = new List<RangeReference>();
        }

        /// <summary>
        ///  Initializes a new instance of the <see cref="CompositeRangeReference" /> class.<br/>
        /// </summary>
        /// <param name="rangeReference">The range reference used to initialise this composite.</param>
        public CompositeRangeReference(RangeReference rangeReference) : this()
        {
            this.Update(rangeReference);
        }

        /// <summary>
        ///  Initializes a new instance of the <see cref="CompositeRangeReference" /> class.<br/>
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public CompositeRangeReference(string sheetName, uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
            : this(new RangeReference(sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex))
        {
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets a text representation of this reference.
        /// </summary>
        public string Reference
        {
            get { return this.reference; }
        }

        /// <summary>
        /// Gets the minimum Excel row index that this<see cref="CompositeRangeReference"/> represents;
        /// </summary>
        public uint MinRowIndex
        {
            get { return this.minRowIndex; }
        }

        /// <summary>
        /// Gets the maximum Excel row index that this<see cref="CompositeRangeReference"/> represents;
        /// </summary>
        public uint MaxRowIndex
        {
            get { return this.maxRowIndex; }
        }

        /// <summary>
        /// Gets the minimum Excel column index that this<see cref="CompositeRangeReference"/> represents;
        /// </summary>
        public uint MinColumnIndex
        {
            get { return this.minColumnIndex; }
        }

        /// <summary>
        /// Gets the maximum Excel column index that this<see cref="CompositeRangeReference"/> represents;
        /// </summary>
        public uint MaxColumnIndex
        {
            get { return this.maxColumnIndex; }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Updates this <see cref="CompositeRangeReference"/> with a new range reference.
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public void Update(string sheetName, uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            this.Update(new RangeReference(sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex));
        }

        /// <summary>
        /// Updates this <see cref="CompositeRangeReference"/> with a new range reference.
        /// </summary>
        /// <param name="rangeReference">The <see cref="RangeReference"/> to be added</param>
        public void Update(RangeReference rangeReference)
        {
            if (this.rangeReferences.Count == 0)
            {
                this.rangeReferences.Add(rangeReference);
                this.sheetName = rangeReference.SheetName;
                this.minRowIndex = rangeReference.StartRowIndex;
                this.maxRowIndex = rangeReference.EndRowIndex;
                this.minColumnIndex = rangeReference.StartColumnIndex;
                this.maxColumnIndex = rangeReference.EndColumnIndex;
            }
            else
            {
                if (this.sheetName != rangeReference.SheetName)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("Cannot create a composite range reference using 2 different worksheets.");
                    sb.AppendFormat("Last added worksheet was '{0}', trying to add '{1}'", this.sheetName, rangeReference.SheetName);

                    throw new InvalidOperationException(sb.ToString());
                }

                RangeReference lrr = this.rangeReferences[this.rangeReferences.Count - 1];

                bool sameColumns = lrr.StartColumnIndex == rangeReference.StartColumnIndex && lrr.EndColumnIndex == rangeReference.EndColumnIndex;
                bool sameRows = lrr.StartRowIndex == rangeReference.StartRowIndex && lrr.EndRowIndex == rangeReference.EndRowIndex;

                if (sameColumns && lrr.EndRowIndex > 0 && lrr.EndRowIndex == (rangeReference.StartRowIndex - 1))
                {
                    // Update last to append after rows
                    lrr.EndRowIndex = rangeReference.EndRowIndex;
                }
                else if (sameColumns && rangeReference.EndRowIndex > 0 && lrr.StartRowIndex == (rangeReference.EndRowIndex + 1))
                {
                    // Update last to append before rows
                    lrr.StartRowIndex = rangeReference.StartRowIndex;
                }
                else if (sameRows && lrr.EndColumnIndex > 0 && lrr.EndColumnIndex == (rangeReference.StartColumnIndex - 1))
                {
                    // Update last to append after columns
                    lrr.EndColumnIndex = rangeReference.EndColumnIndex;
                }
                else if (sameColumns && rangeReference.EndRowIndex > 0 && lrr.StartRowIndex == (rangeReference.EndRowIndex + 1))
                {
                    // Update last to append before columns
                    lrr.StartColumnIndex = rangeReference.StartColumnIndex;
                }
                else if (!sameColumns || !sameRows)
                {
                    // Finally, unless we actually have the same range, can't append, so add to list...
                    this.rangeReferences.Add(rangeReference);
                }

                // Update min/max row and column indexes
                if (rangeReference.StartRowIndex < this.minRowIndex) this.minRowIndex = rangeReference.StartRowIndex;
                if (rangeReference.EndRowIndex > this.maxRowIndex) this.maxRowIndex = rangeReference.EndRowIndex;
                if (rangeReference.StartColumnIndex < this.minColumnIndex) this.minColumnIndex = rangeReference.StartColumnIndex;
                if (rangeReference.EndColumnIndex > this.maxColumnIndex) this.maxColumnIndex = rangeReference.EndRowIndex;

            }

            this.reference = this.GetExcelAbsoluteRef();
        }

        /// <summary>
        /// Gets a string representation of this object instance.
        /// </summary>
        /// <returns>String representation of this object instance.</returns>
        public override string ToString()
        {
            return string.Format("{0}:{1}", base.ToString(), this.reference);
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Returns a comma separated $row$column reference to this list of <see cref="RangeReference"/>s
        /// </summary>
        /// <returns>A comma separated $row$column reference to this list of <see cref="RangeReference"/>s</returns>
        private string GetExcelAbsoluteRef()
        {
            var sb = new StringBuilder();
            foreach (RangeReference rr in this.rangeReferences)
            {
                sb.AppendFormat("{0},", rr.GetExcelAbsoluteRef());
            }

            if (sb.Length > 0)
            {
                return sb.ToString().Remove(sb.Length - 1);
            }

            return null;
        }

        #endregion Private Helpers
    }
}

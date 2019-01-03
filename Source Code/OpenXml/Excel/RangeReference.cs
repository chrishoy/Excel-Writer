// -----------------------------------------------------------------------
// <copyright file="RangeReference.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a range reference in Excel
    /// </summary>
    public class RangeReference
    {
        public string SheetName { get; set; }
        public uint StartRowIndex { get; set; }
        public uint StartColumnIndex { get; set; }
        public uint EndRowIndex { get; set; }
        public uint EndColumnIndex { get; set; }

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public RangeReference(uint rowIndex, uint columnIndex)
            : this(null, rowIndex, columnIndex, rowIndex, columnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public RangeReference(string sheetName, uint rowIndex, uint columnIndex)
            : this(sheetName, rowIndex, columnIndex, rowIndex, columnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="rowIndex">The row index</param>
        /// <param name="columnIndex">The column index</param>
        public RangeReference(int rowIndex, int columnIndex)
            : this(null, (uint)rowIndex, (uint)columnIndex, (uint)rowIndex, (uint)columnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="rowIndex">The row index</param>
        /// <param name="columnIndex">The column index</param>
        public RangeReference(string sheetName, int rowIndex, int columnIndex)
            : this(sheetName, (uint)rowIndex, (uint)columnIndex, (uint)rowIndex, (uint)columnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public RangeReference(uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
            : this(null, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public RangeReference(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
            : this(null, (uint)startRowIndex, (uint)startColumnIndex, (uint)endRowIndex, (uint)endColumnIndex)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeReference" /> class
        /// </summary>
        /// <param name="sheetName">The worksheet name</param>
        /// <param name="startRowIndex">The start row index</param>
        /// <param name="startColumnIndex">The start column index</param>
        /// <param name="endRowIndex">The end row index</param>
        /// <param name="endColumnIndex">The end column index</param>
        public RangeReference(string sheetName, uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            this.SheetName = sheetName;
            this.StartRowIndex = startRowIndex;
            this.StartColumnIndex = startColumnIndex;
            this.EndRowIndex = endRowIndex;
            this.EndColumnIndex = endColumnIndex;
        }

        #endregion Construction

        /// <summary>
        /// Returns a $row$column reference to this <see cref="RangeReference"/>
        /// </summary>
        /// <returns></returns>
        public string GetExcelAbsoluteRef()
        {
            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                if (string.IsNullOrEmpty(this.SheetName))
                {
                    return GetCellRef(this.StartRowIndex, this.StartColumnIndex, true);
                }
                else
                {
                    return string.Format("'{0}'!{1}", this.SheetName, GetCellRef(this.StartRowIndex, this.StartColumnIndex, true));
                }
            }
            if (string.IsNullOrEmpty(this.SheetName))
            {
                return string.Format("{0}:{1}", GetCellRef(this.StartRowIndex, this.StartColumnIndex, true)
                                , GetCellRef(this.EndRowIndex, this.EndColumnIndex, true));
            }
            else
            {
                return string.Format("'{0}'!{1}:{2}", this.SheetName, GetCellRef(this.StartRowIndex, this.StartColumnIndex, true)
                                , GetCellRef(this.EndRowIndex, this.EndColumnIndex, true));
            }
        }

        private static string GetCellRef(uint rowIndex, uint columnIndex, bool absolute)
        {
            if (absolute)
            {
                return string.Format("${0}${1}", CellExtensions.GetColumnLetter(columnIndex), rowIndex);
            }
            else
            {
                return string.Format("{0}{1}", CellExtensions.GetColumnLetter(columnIndex), rowIndex);
            }
        }

        public override string ToString()
        {
            return string.Format("{0}={1}", base.ToString(), this.GetExcelAbsoluteRef());
        }
    }
}

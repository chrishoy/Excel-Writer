// -----------------------------------------------------------------------
// <copyright file="TableDataRowInfo.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents information relating to a row as it is created for TableData.<br/>
    /// This information is later used to process charts.
    /// </summary>
    internal class TableDataRowInfo
    {
        #region Private Fields

        private object rowData;
        private uint tableRowIndex;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor. Requires <see cref="object">rowData</see>.
        /// </summary>
        /// <param name="rowData"></param>
        public TableDataRowInfo(object rowData)
        {
            this.rowData = rowData;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Represents the data which is supplied to a row of <see cref="TableData"/>
        /// </summary>
        public object RowData
        {
            get { return this.rowData; }
            set { this.rowData = value; }
        }

        /// <summary>
        /// Row index in <see cref="Table"/> (to be written into Excel worksheet), where the <see cref="TableData"/> row resides.<br/>
        /// Note that this is not the Excel row index, but the mapped row index within the <see cref="Table"/>.
        /// </summary>
        public uint TableRowIndex
        {
            get { return this.tableRowIndex; }
            set { this.tableRowIndex = value; }
        }

        #endregion Public Properties

    }
}

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows;
    using System.Windows.Media;

    /// <summary>
    /// Maintains information about a cell that is to be written into an excel worksheet.<br/>
    /// This includes style information, the value that is to be written, and row/column spanning information.
    /// </summary>
    internal class ExcelCellInfo
    {
        #region Private Fields

        private ExcelCellStyleInfo styleInfo;
        private object value;

        private uint lastSpanColumn;
        private uint lastSpanRow;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Ctor. Initialises an empty instance of a <see cref="ExcelCellInfo"/>
        /// </summary>
        public ExcelCellInfo()
        {
            this.styleInfo = new ExcelCellStyleInfo();
            this.value = string.Empty;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Returns the style related information used to format Cell in an excel worksheet
        /// </summary>
        public ExcelCellStyleInfo StyleInfo
        {
            get { return this.styleInfo; }
            set { this.styleInfo = value; }
        }

        /// <summary>
        /// Get/set the value that is to be written into the Excel cell.
        /// </summary>
        public object Value
        {
            get { return this.value; }
            set { this.value = value; }
        }

        /// <summary>
        /// Returns true if there is either style information or a value associated with this cell.
        /// </summary>
        public bool HasStyleOrValue
        {
            get { return this.styleInfo.HasCellInfo || this.value != null;  }
        }

        /// <summary>
        /// Get/set whether this cell information is to span to last column of<br/>
        /// the outmost containing entity (top layer) when exported to Excel.
        /// </summary>
        public uint LastSpanColumn
        {
            get { return this.lastSpanColumn; }
            set { this.lastSpanColumn = value; }
        }

        /// <summary>
        /// Get/set whether this cell information is to span to last row of<br/>
        /// the outmost containing entity (top layer) when exported to Excel.
        /// </summary>
        public uint LastSpanRow
        {
            get { return this.lastSpanRow; }
            set { this.lastSpanRow = value; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets or sets the cell that this cell is to be merged with.
        /// </summary>
        internal ExcelCellInfo MergeFrom { get; set; }

        /// <summary>
        /// Gets or sets the range of cells that this cell is to be merged into.
        /// </summary>
        internal ExcelCellInfo MergeTo { get; set; }

        /// <summary>
        /// Gets or sets the cell created.
        /// </summary>
        internal DocumentFormat.OpenXml.Spreadsheet.Cell Cell { get; set;}

        #endregion Internal Properties
    }
}

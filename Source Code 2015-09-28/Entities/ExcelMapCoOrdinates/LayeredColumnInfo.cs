namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents the layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that will have to be processed
    /// to determine what is to be written into an Excel column.
    /// </summary>
    internal class LayeredColumnInfo
    {
        #region Construction

        /// <summary>
        /// asd
        /// </summary>
        public LayeredColumnInfo()
        {
            this.Maps = new List<ExcelMapCoOrdinate>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the set of layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that 
        /// will have to be processed when determining what is to be written into a column in Excel.
        /// </summary>
        public List<ExcelMapCoOrdinate> Maps { get; set; }

        /// <summary>
        /// Gets or sets the column information (formatting and size) that is to be written into a single column in Excel.<br/>
        /// This is the result of processing the Maps property.
        /// </summary>
        public ExcelColumnInfo ColumnInfo { get; set; }

        #endregion Public Properties
    }
}

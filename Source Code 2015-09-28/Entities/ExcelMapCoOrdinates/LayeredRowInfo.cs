namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents the layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that will have to be processed
    /// to determine what is to be written into an Excel row.
    /// </summary>
    internal class LayeredRowInfo
    {
        #region Construction

        /// <summary>
        /// asd
        /// </summary>
        public LayeredRowInfo()
        {
            this.Maps = new List<ExcelMapCoOrdinate>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the set of layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that 
        /// will have to be processed when determining what is to be written into a row in Excel.
        /// </summary>
        public List<ExcelMapCoOrdinate> Maps { get; set; }

        /// <summary>
        /// Gets or sets the row information (formatting and size) that is to be written into a single row in Excel.<br/>
        /// This is the result of processing the Maps property.
        /// </summary>
        public ExcelRowInfo RowInfo { get; set; }

        #endregion Public Properties
    }
}

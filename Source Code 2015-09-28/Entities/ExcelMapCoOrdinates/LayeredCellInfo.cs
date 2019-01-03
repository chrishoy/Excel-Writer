namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents the layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that will have to be processed
    /// to determine what is to be written into an Excel Cell.
    /// </summary>
    internal class LayeredCellInfo
    {
        #region Construction

        /// <summary>
        /// asd
        /// </summary>
        public LayeredCellInfo()
        {
            this.LayeredMaps = new List<ExcelMapCoOrdinate>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the set of layered <see cref="ExcelMapCoOrdinate">Containers and Cells</see> that 
        /// will have to be processed when determining what is to be written into a single cell in Excel.
        /// </summary>
        public List<ExcelMapCoOrdinate> LayeredMaps { get; set; }

        /// <summary>
        /// Gets or sets the cell information (formatting and value) that is to be written into a single cell in Exce.<br/>
        /// This is the result of processing the LayeredMaps property.
        /// </summary>
        public ExcelCellInfo CellInfo { get; set; }

        #endregion Public Properties
    }
}

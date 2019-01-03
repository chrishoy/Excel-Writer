using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Each IDataPart requires a PartId and Data
    /// The PartId is needed to link together IDataPart with IPart (export and report parts).
    /// Once the linked the data is transferred from IDataPart to the IPart
    /// </summary>
    public interface IDataPart
    {
        /// <summary>
        /// The unique identifier of the data part
        /// </summary>
        string PartId { get; set; }

        /// <summary>
        /// The data associated with the data part
        /// </summary>
        object Data { get; }

        /// <summary>
        /// The number or rows in the 'table' bit of this part
        /// For example when this data part is associated with an export view item
        /// this will be the number of items in the itemsource property
        /// </summary>
        int RowCount { get; }
    }
}

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Defines the properties of a <see cref="BaseMap"/> derived class which
    /// can be manipulated during a Prepare operation on a <see cref="IPreparable"/> data part.
    /// </summary>
    public interface IExcelPreparable
    {
    }

    /// <summary>
    /// Defines the properties of a <see cref="TableData"/> derived class which
    /// can be manipulated during a Prepare operation on a <see cref="IPreparable"/> data part.
    /// </summary>
    public interface IExcelTableDataPreparable : IExcelPreparable
    {
        /// <summary>
        /// Gets a collection of <see cref="TableColumn"/>s
        /// </summary>
        TableColumnCollection Columns { get; }
    }

    /// <summary>
    /// Defines the properties of a <see cref="Table"/> derived class which
    /// can be manipulated during a Prepare operation on a <see cref="IPreparable"/> data part.
    /// </summary>
    public interface IExcelTablePreparable : IExcelPreparable
    {
        /// <summary>
        /// Gets a collection of <see cref="TableColumn"/>s
        /// </summary>
        TableColumnCollection Columns { get; }

        /// <summary>
        /// Gets a collection of <see cref="TableColumnHeader"/>s (span across <see cref="TableColumns"/>s).
        /// </summary>
        TableColumnHeaderCollection ColumnHeaders { get; }
    }
}

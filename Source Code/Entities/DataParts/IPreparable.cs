using System.Windows;

namespace ExcelWriter
{
    /// <summary>
    /// For dynamically changing data
    /// </summary>
    public interface IPreparable
    {
        /// <summary>
        /// Required to dynamically modify an export element, such as a cell or table
        /// based on the contents of the data. The element being modified is the elementToPrepare.
        /// </summary>
        /// <param name="parent">Not currently used</param>
        /// <param name="elementToPrepare">The element being modified to support the data.</param>
        void Prepare(object parent, IExcelPreparable elementToPrepare);
    }
}

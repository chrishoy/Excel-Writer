
namespace ExcelWriter
{
    /// <summary>
    /// Used to select style using values in the supplied item
    /// </summary>
    public abstract class CellStyleSelector : IResource
    {
        /// <summary>
        /// Get/set the key for this selector within a dictionary.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Contains logic for deciding which cell style key to return.
        /// The item passed in (DataContext) can be used in the logic.
        /// </summary>
        public abstract string SelectCellStyleKey(object item);
    }
}

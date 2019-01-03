namespace ExcelWriter
{
    /// <summary>
    /// Represents an item that can be stored in a <see cref="Store"/>
    /// </summary>
    public interface IStorable
    {
        /// <summary>
        /// Gets the ID of the item in the store
        /// </summary>
        int Id { get; }
    }
}

namespace ExcelWriter
{
    /// <summary>
    /// Any object used as a Resource much implememt this interface.
    /// These are Styles, StyleSelectors, Maps, TemplateCollections
    /// </summary>
    public interface IResource
    {
        /// <summary>
        /// Gets or sets the key which identifies a resource.
        /// </summary>
        string Key { get; set; }
    }
}

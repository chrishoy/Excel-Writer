namespace ExcelWriter
{
    /// <summary>
    /// A class that holds resources must implement this interface
    /// </summary>
    public interface IResourceContainer
    {
        string DesignerFileName { get; set; }
        ResourceCollection Resources { get; set; }
    }
}

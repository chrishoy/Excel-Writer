namespace ExcelWriter
{
    using System.Collections.Generic;

    /// <summary>
    /// Defines export files for the export map service.
    /// TODO: Consumer does not need to know about this class... Find a better way to hide it...!
    /// </summary>
    public class ExportFiles
    {
        /// <summary>
        /// Initialises a new instance of the <see cref="ExportFiles"/> class.
        /// </summary>
        public ExportFiles()
        {
            this.Metadata = new Dictionary<string, string>();
            this.Templates = new Dictionary<string, byte[]>();
        }

        /// <summary>
        /// Gets or sets the metadata.
        /// </summary>
        public Dictionary<string, string> Metadata { get; set; }

        /// <summary>
        /// Gets or sets the templates.
        /// </summary>
        public Dictionary<string, byte[]> Templates { get; set; }
    }
}

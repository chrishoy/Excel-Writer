namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Windows.Markup;
    using System.Xml;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Represents a collection of 
    /// </summary>
    public sealed class Book : DocumentMetadataBase

    {
        public Book()
        {
            this.Parts = new List<ExportPart>();
            this.MappingPlaceholders = new List<MappingPlaceholder>();
            this.MappingPlaceholderSets = new List<MappingPlaceholderSet>();
        }

        /// <summary>
        /// The identifier of this metadata
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// A list of placeholder which are used as anchors to move shapes from sheet to sheet
        /// This allows for 
        /// </summary>
        public List<MappingPlaceholder> MappingPlaceholders { get; set; }

        /// <summary>
        /// </summary>
        public List<MappingPlaceholderSet> MappingPlaceholderSets { get; set; }

        /// <summary>
        /// The list of parts
        /// </summary>
        public List<ExportPart> Parts { get; set; }

        public override IEnumerable<string> GetPartIds()
        {
            if (this.Parts == null) 
            {
                return new List<string>();
            }

            return (from p in this.Parts
                    where !string.IsNullOrEmpty(p.PartId)
                    select p.PartId).Distinct();
        }         

        public bool Validate(List<string> errors)
        {
            // For each part 
            return true;
        }
    }
}

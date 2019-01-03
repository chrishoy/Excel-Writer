using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// The ExportPart together with the Template, which it is linked using the TemplateId, 
    /// provides information on pulling data into an excel workbook.
    /// Each part has a data sheet, which contains the raw data associated with the part, 
    /// and a user sheet, which has the formatted data in the style of a chart, table or any excel construct.
    /// </summary>
    public sealed class ExportPart
    {
        public ExportPart()
        {
            this.CompositeTemplateMappings = new List<TemplateMapping>();
            this.Mappings = new List<VisualMapping>();
        }

        /// <summary>
        /// </summary>
        public List<TemplateMapping> CompositeTemplateMappings { get; set; }

        /// <summary>
        /// When true the data sheet is hidden
        /// </summary>
        public bool DataSheetHidden { get; set; }

        /// <summary>
        /// The name of the raw data sheet to create, this is a cloned from the template
        /// If a title is specified on the template this will be used instead
        /// </summary>
        public string DataSheetName { get; set; }

        /// <summary>
        /// If a title is specified on the template this will be used instead
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// If true, then the partId-TemplateId mapping will take place against the
        /// CompositeDataPart DataParts collection
        /// </summary>
        public bool IsComposite { get; set; }

        /// <summary>
        /// Will throw an export exception a mandatory export part is not 
        /// provided with a data part during export
        /// </summary>
        public bool IsMandatory { get; set; }

        /// <summary>
        /// A list of instructions on how to map a drawing or range of cells
        /// </summary>
        public List<VisualMapping> Mappings { get; set; }

        /// <summary>
        /// Unique part id
        /// </summary>
        public string PartId { get; set; }

        /// <summary>
        /// The name of the presentation sheet to create, this is a cloned from the template
        /// If a title is specified on the template this will be used instead
        /// The presentation sheet will generally contain charts and other drawing as well 
        /// as formatted tables that could be used in report
        /// </summary>
        public string PresentationSheetName { get; set; }

        /// <summary>
        /// Presentation sheets are removed by default if any mappings are provided.
        /// This suppresses that behaviour.
        /// </summary>
        public bool SuppressPresentationRemoval { get; set; }

        /// <summary>
        /// The id of the template associated with this part
        /// </summary>
        public string TemplateId { get; set; }

        /// <summary>
        /// </summary>
        public TemplateMapping TemplateMapping { get; set; }

        /// <summary>
        /// PartId mandatory
        /// TemplatePartId mandatory
        /// </summary>
        public bool Valid(out string error)
        {
            error = null;

            StringBuilder s = new StringBuilder();

            if (string.IsNullOrEmpty(this.PartId))
            {
                s.Append("No PartId specified");
                s.Append(Environment.NewLine);
            }
            if (string.IsNullOrEmpty(this.TemplateId))
            {
                s.Append("No TemplateId specified");
                s.Append(Environment.NewLine);
            }

            if (s.Length > 0)
            {
                error = s.ToString();
                return false;
            }
            return true;
        }
    }
}

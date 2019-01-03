namespace ExcelWriter
{
    using System.Collections.Generic;

    /// <summary>
    /// The export map service. This service wraps the MM.Framework.Export.Map class, which generated Excel documents by mapping data to Excel using XAML.
    /// </summary>
    public interface IExportMapService
    {
        /// <summary>
        /// The generate report method.  This will take the supplied information and use it to generate a report.
        /// </summary>
        /// <param name="exportDataParts">The data which will be rendered into report parts.</param>
        /// <param name="exportMetadataFiles"> The metadata files. </param>
        /// <param name="exportTemplateFiles"> The template files. </param>
        /// <param name="isDocumentMetaBased"> The is document meta based. </param>
        /// <param name="isPdf">If true, the returned report will be a PDF, otherwise it will be an Excel document.</param>
        /// <returns>A stream which contains the report</returns>
        MemoryStreamResult GenerateReport(
            List<ExportDataPart> exportDataParts, 
            ExportFiles exportMetadataFiles, 
            ExportFiles exportTemplateFiles, 
            bool isDocumentMetaBased = true, 
            bool isPdf = false);
    }
}

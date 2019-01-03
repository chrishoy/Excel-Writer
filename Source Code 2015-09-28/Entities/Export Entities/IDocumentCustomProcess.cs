namespace ExcelWriter
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Implement if you want to perform custom pre and post processing during document generation
    /// </summary>
    public interface IDocumentCustomProcess
    {
        /// <summary>
        /// Hook for performing custom pre processing steps on the document
        /// </summary>
        /// <param name="document">The OpenXml SpreadSheetdocument.</param>
        /// <param name="exportParameters">The export parameters useful to know if the document is going to be PDFd or not.</param>
        /// <param name="metadata">The XAML metadata used to drive creation of the document.</param>
        /// <param name="dataParts">The collection of data parts bound to the XAML metadata</param>        
        void PreProcess(SpreadsheetDocument document,
                        ExportParameters exportParameters, 
                        DocumentMetadataBase metadata, 
                        IEnumerable<IDataPart> dataParts);

        /// <summary>
        /// Hook for performing custom post processing steps on the document
        /// </summary>
        /// <param name="document">The OpenXml SpreadSheetdocument.</param>
        /// <param name="exportParameters">The export parameters useful to know if the document is going to be PDFd or not.</param>
        /// <param name="metadata">The XAML metadata used to drive creation of the document.</param>
        /// <param name="dataParts">The collection of data parts bound to the XAML metadata</param>
        void PostProcess(SpreadsheetDocument document,
                         ExportParameters exportParameters,
                         DocumentMetadataBase metadata, 
                         IEnumerable<IDataPart> dataParts);
    }
}

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    ///
    /// </summary>
    public sealed partial class ExportGenerator
    {
        ///// <summary>
        ///// Creates an excel file containing data provided in the dataparts.
        ///// Each IDatapart will match with an ExportPart of the same PartId
        ///// and will use this information to output the data in the desired format.
        ///// </summary>
        ///// <param name="open">true to open, false to not</param>
        ///// <param name="title">The title.</param>
        ///// <param name="dataParts">The data parts to output</param>
        ///// <param name="metadata">The metadata.</param>
        ///// <param name="exportParameters">Optional export parameters.</param>
        ///// <returns>The full path of the created file</returns>        
        //public string ExportToExcel(bool open, string title, IEnumerable<IDataPart> dataParts, ExportMetadata metadata, ExcelTemplatePackage templatePackage, ExportParameters exportParameters = null)
        //{
        //    if (string.IsNullOrEmpty(title))
        //    {
        //        throw new ArgumentNullException("targetFilePath");
        //    }
        //    if (dataParts == null)
        //    {
        //        throw new ArgumentNullException("dataParts");
        //    }
        //    if (metadata == null)
        //    {
        //        throw new ArgumentNullException("metadata");
        //    }
        //    if (templatePackage == null)
        //    {
        //        throw new ArgumentNullException("templatePackage");
        //    }

        //    // Adds the parameters as a data part (creates a hidden sheet)
        //    dataParts = AddDebugPart(dataParts, metadata, exportParameters);

        //    // build our sets of DataParts, ExportParts, ExportTemplates
        //    List<ExportTripleSet> sets = this.BuildSets(dataParts, metadata, templatePackage);
        //    if (sets.Count == 0)
        //    {
        //        throw new ExportException("Nothing to export, no matching DataParts, ExportParts, ExportTemplates found");
        //    }

        //    var stream = this.ExcelExportInternal(exportParameters, metadata, sets, dataParts, templatePackage);

        //    if (!title.EndsWith(".xls") && !title.EndsWith(".xlsx"))
        //    {
        //        title = string.Concat(title, ".xlsx");
        //    }

        //    return SaveAndOpen(open, title, stream);
        //}

        /// <summary>
        /// Exports to excel memory stream.
        /// </summary>
        /// <param name="dataParts">The data parts.</param>
        /// <param name="metadata">The metadata.</param>
        /// <param name="templatePackage">The template package.</param>
        /// <param name="exportParameters">The export parameters.</param>
        /// <returns></returns>
        public ExportToMemoryStreamResult ExportToExcelMemoryStream(
            IEnumerable<IDataPart> dataParts, 
            Book metadata, 
            ExcelTemplatePackage resourcePackage, 
            ExportParameters exportParameters = null)
        {
            Guard.IsNotNull(dataParts, "dataParts");
            Guard.IsNotNull(metadata, "metadata");
            Guard.IsNotNull(resourcePackage, "resourcePackage");

            var result = new ExportToMemoryStreamResult();
      
            dataParts = AddDebugPart(dataParts, metadata, exportParameters);

            // build our sets of DataParts, ExportParts, ExportTemplates
            List<ExportTripleSet> sets = this.BuildSets(dataParts, metadata, resourcePackage);
            if (sets.Count == 0)
            {
                throw new ExportException("Nothing to export, no matching DataParts, ExportParts, ExportTemplates found");
            }

            try
            {
                result.MemoryStream = this.ExcelExportInternal(exportParameters, metadata, sets, dataParts, resourcePackage);
            }
            catch (Exception ex)
            {
                result.Error = ex;
            }

            return result;
        }

        /// <summary>
        /// Generates the document.
        /// </summary>
        /// <param name="dataParts">The data parts.</param>
        /// <param name="metadata">The metadata.</param>
        /// <param name="resourcePackage">The resource package.</param>
        /// <param name="exportParameters">The export parameters.</param>
        /// <returns></returns>
        public ExportToMemoryStreamResult GenerateDocument(IEnumerable<IDataPart> dataParts, DocumentMetadataBase metadata, ResourcePackage resourcePackage, ExportParameters exportParameters = null)
        {
            Guard.IsNotNull(dataParts, "dataParts");
            Guard.IsNotNull(metadata, "metadata");
            Guard.IsNotNull(resourcePackage, "resourcePackage");

            var result = new ExportToMemoryStreamResult();

            try
            {
                if (metadata.DocumentMetadataType == DocumentMetadataType.ExcelDocument)
                {
                    result.MemoryStream = this.GenerateExcelInternal(exportParameters, (ExcelDocumentMetadata)metadata, dataParts, resourcePackage);
                }
                else
                {
                    throw new ExportException("Unknown document metadata type");
                }
            }
            catch (Exception ex)
            {
                result.Error = ex;
            }
            return result;
        }
    }
}

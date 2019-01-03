namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    using Interface;

    /// <summary>
    /// The export map service. This service wraps the MM.Framework.Export.Map class
    /// </summary>
    public class ExportMapService : IExportMapService
    {
        #region Private fields

        private ILogger logger;

        private bool isDocumentMetaBased;
        private ExportFiles metadataFiles;
        private ExportFiles templateFiles;
        private List<IDataPart> dataParts; 

        private string tempSubFolder;
        private string sourceFolder;
        private string metadataFolder;
        private string templateFolder;
        private string metadataPackageName;
        private string templatePackageName;

        #endregion Private fields

        #region Construction

        public ExportMapService(ILogger logger)
        {
            Guard.IsNotNull(logger, "logger");
            this.logger = logger;
        }

        public ExportMapService()
        {
            this.logger = new NullLogger();
        }

        #endregion Construction

        #region Public methods

        /// <summary>
        /// The generate report method.  This will take the supplied information and use it to generate a report.
        /// </summary>
        /// <param name="exportDataParts">The data which will be rendered into report parts.</param>
        /// <param name="exportMetadataFiles"> The metadata files. </param>
        /// <param name="exportTemplateFiles"> The template files. </param>
        /// <param name="isDocumentMetadataBased"> The is document metadata based. </param>
        /// <param name="isPdf">If true, the returned report will be a PDF, otherwise it will be an Excel document.</param>
        /// <returns>A stream which contains the report</returns>
        public MemoryStreamResult GenerateReport(List<ExportDataPart> exportDataParts, ExportFiles exportMetadataFiles, ExportFiles exportTemplateFiles, bool isDocumentMetadataBased = true, bool isPdf = false)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "Enter {0}.GenerateReport", this.GetType().Name);

            Guard.IsNotNull(exportDataParts, "exportDataParts");
            Guard.IsNotNull(exportMetadataFiles, "exportMetadataFiles");
            Guard.IsNotNull(exportTemplateFiles, "exportTemplateFiles");

            this.dataParts = exportDataParts.ConvertAll(dp => new GenericExportDataPart(dp) as IDataPart);
            this.metadataFiles = exportMetadataFiles;
            this.templateFiles = exportTemplateFiles;
            this.isDocumentMetaBased = isDocumentMetadataBased;
            this.tempSubFolder = Path.GetRandomFileName();
            
            // Write supplied files to disk before we pack them.
            this.WriteExportFiles();

            // Pack the report metadata and template files
            this.metadataPackageName = this.PackSourceMetaDataFolder(this.metadataFolder);
            this.templatePackageName = this.PackSourceTemplatesFolder(this.templateFolder);

            MemoryStreamResult result;

            // Generate the report
            if (this.isDocumentMetaBased)
            {
                // Version 1 way of generating a report
                result = this.GenerateDocumentMetadataBasedReport(isPdf);
            }
            else
            {
                // Version 2 way of generating a report (as far as I know, not yet completed)
                result = this.GenerateExportMetadataReport(isPdf);
            }

            this.DeleteExportFiles();

            // Reset memory stream to 0 position for return.
            if (result != null && result.MemoryStream != null)
            {
                result.MemoryStream.Position = 0;
            }

            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "Exit {0}.GenerateReport", this.GetType().Name);

            return result;
        }

        #endregion Public methods

        #region Private methods

        /// <summary>
        /// The generate document metadata based report.
        /// </summary>
        private MemoryStreamResult GenerateDocumentMetadataBasedReport(bool isPdf)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.GenerateDocumentMetadataBasedReport", this.GetType().Name);
            var resourcesPackage = ResourcePackage.Open(Path.Combine(this.sourceFolder, this.templatePackageName));
            var documentMetadataPackage = DocumentMetadataPackage.Open(Path.Combine(this.sourceFolder, this.metadataPackageName)).DocumentMetadata;
            var generator = new ExportGenerator();
            var exportParams = new ExportParameters { ConvertOutputToPdf = isPdf };

            ExportToMemoryStreamResult result = generator.GenerateDocument(this.dataParts,
                                                                           documentMetadataPackage,
                                                                           resourcesPackage,
                                                                           exportParams);

            if (result.Error != null)
            {
                this.logger.Log(LogType.Fatal, this.GetAssemblyName(), "{0}.GenerateDocumentMetadataBasedReport", result.Error);
            }

            return new MemoryStreamResult
            {
                MemoryStream = result.MemoryStream,
                Status = result.Error == null ? MemoryStreamResultStatus.Success : MemoryStreamResultStatus.Failure,
                ErrorMessage = result.Error == null ? null : result.Error.Message,
            };
        }

        /// <summary>
        /// The generate export metadata report.
        /// </summary>
        /// <param name="isPdf">The is PDF.</param>
        private MemoryStreamResult GenerateExportMetadataReport(bool isPdf)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.GenerateExportMetadataReport", this.GetType().Name);
            var exceltemplatePackage = ExcelTemplatePackage.Open(Path.Combine(this.sourceFolder, this.templatePackageName));
            var exportMetadata = ExportMetadataPackage.Open(Path.Combine(this.sourceFolder, this.metadataPackageName)).ExportMetadata;
            var generator = new ExportGenerator();
            var exportParams = new ExportParameters { ConvertOutputToPdf = isPdf };

            ExportToMemoryStreamResult result = generator.ExportToExcelMemoryStream(this.dataParts,
                                                                                    exportMetadata,
                                                                                    exceltemplatePackage,
                                                                                    exportParams);
            if (result.Error != null)
            {
                this.logger.Log(LogType.Fatal, this.GetAssemblyName(), "{0}.GenerateDocumentMetadataBasedReport", result.Error);
            }
            
            return new MemoryStreamResult
            {
                MemoryStream = result.MemoryStream,
                Status = result.Error == null ? MemoryStreamResultStatus.Success : MemoryStreamResultStatus.Failure,
                ErrorMessage = result.Error == null ? null : result.Error.Message,
            };
        }

        /// <summary>
        /// The pack source data folder.
        /// </summary>
        /// <param name="subFolderName">The sub folder name.</param>
        /// <returns>The zipped file name</returns>
        private string PackSourceMetaDataFolder(string subFolderName)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.PackSourceMetaDataFolder - Packing {1}", this.GetType().Name, subFolderName);
            string directory = Path.Combine(this.sourceFolder, subFolderName);
            string package = this.isDocumentMetaBased ? DocumentMetadataPackage.Pack(directory) : ExportMetadataPackage.Pack(directory);

            // Extract package name from returned
            return Path.GetFileName(package);
        }

        /// <summary>
        /// The pack source templates folder.
        /// </summary>
        /// <param name="subFolderName">The sub folder name.</param>
        /// <returns>The zipped file name</returns>
        private string PackSourceTemplatesFolder(string subFolderName)
        {
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.PackSourceTemplatesFolder - Packing {1}", this.GetType().Name, subFolderName);
            string directory = Path.Combine(this.sourceFolder, subFolderName);
            string package = this.isDocumentMetaBased ? ResourcePackage.Pack(directory) : ExcelTemplatePackage.Pack(directory);

            // Extract package name from returned
            return Path.GetFileName(package);
        }

        /// <summary>
        /// The write export files.
        /// </summary>
        private void WriteExportFiles()
        {
            // Set the folders the generate code expects
            this.sourceFolder = Path.Combine(Path.GetTempPath(), this.tempSubFolder);
            this.metadataFolder = "Metadata";
            this.templateFolder = "Template";

            // Get the full paths to the metadata and template folders
            var metadataPath = Path.Combine(this.sourceFolder, this.metadataFolder);
            var templatePath = Path.Combine(this.sourceFolder, this.templateFolder);

            // Ensure all the folders exist
            this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.WriteExportFiles - Creating Temp Folders {1}", this.GetType().Name, this.sourceFolder);
            Directory.CreateDirectory(this.sourceFolder);
            Directory.CreateDirectory(metadataPath);
            Directory.CreateDirectory(templatePath);

            // Write the supplied export files

            // Write the metadata files (NB! non-document metadata only has one template file in a TemplateFile folder)
            this.WriteExportTextFiles(this.metadataFiles, metadataPath, "Metadata");
            this.WriteExportByteFiles(this.metadataFiles, metadataPath, this.isDocumentMetaBased ? "TemplateFiles" : "TemplateFile");

            // Write the template files
            this.WriteExportTextFiles(this.templateFiles, templatePath, "Metadata");
            this.WriteExportByteFiles(this.templateFiles, templatePath, "TemplateFiles");
        }

        /// <summary>
        /// Writes the supplied export files to folders below the supplied folder.
        /// </summary>
        /// <param name="files"> The files to write. </param>
        /// <param name="folderName"> The folder to write to. </param>
        private void WriteExportTextFiles(ExportFiles files, string folderName, string subFolderName)
        {
            // Create the metadata folder and write files to it
            var folder = Path.Combine(folderName, subFolderName);
            Directory.CreateDirectory(folder);
            foreach (var file in files.Metadata)
            {
                File.WriteAllText(Path.Combine(folder, file.Key), file.Value);
            }
        }

        /// <summary>
        /// Writes the supplied export files to folders below the supplied folder.
        /// </summary>
        /// <param name="files"> The files to write. </param>
        /// <param name="folderName"> The folder to write to.</param>
        /// <param name="subFolderName">The sub-folder to write to.</param>
        private void WriteExportByteFiles(ExportFiles files, string folderName, string subFolderName)
        {
            // Create the template folder and write files to it
            var folder = Path.Combine(folderName, subFolderName);
            Directory.CreateDirectory(folder);
            foreach (var file in files.Templates)
            {
                File.WriteAllBytes(Path.Combine(folder, file.Key), file.Value);
            }
        }

        /// <summary>
        /// The delete export files.
        /// </summary>
        private void DeleteExportFiles()
        {
            try
            {
                if (Directory.Exists(this.sourceFolder))
                {
                    this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.DeleteExportFiles - {1}", this.GetType().Name, this.sourceFolder); 
                    Directory.Delete(this.sourceFolder, true);
                }
            }
            catch (Exception ex)
            {
                // We don't care about errors at this stage
                this.logger.Log(LogType.Trace, this.GetAssemblyName(), "{0}.DeleteExportFiles - FAILED: {1}", this.GetType().Name, ex.Message); 
            }
        }

        #endregion Private methods
    }
}

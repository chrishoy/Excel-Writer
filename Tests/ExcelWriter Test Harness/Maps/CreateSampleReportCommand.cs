namespace ExcelWriter.TestHarness.Maps
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;

    using Interface;
    using ExcelWriter;

    public enum CreationType
    {
        ExcelMetadataBased,
        ExportMetadataBased,
    }

    /// <summary>
    /// Command which, when executed, creates a position report.
    /// </summary>
    public class CreateSampleReportCommand
    {
        private readonly ILogger _logger;
        private readonly IExportMapService _exportMapService;
        private readonly bool _debuggerIsAttached;

        /// <summary>
        /// Initialises a new instance of the <see cref="CreateSampleReportCommand"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="exportMapService">The export map service.</param>
        public CreateSampleReportCommand(
            ILogger logger,
            IExportMapService exportMapService)

        {
            Guard.IsNotNull(logger, "logger");
            Guard.IsNotNull(exportMapService, "exportMapService");

            _logger = logger;
            _exportMapService = exportMapService;
            _debuggerIsAttached = Debugger.IsAttached;
        }

        public MemoryStreamResult Execute(object sourceData, CreationType exportCreationType)
        {
            _logger.Log(LogType.Trace, this.GetAssemblyName(), string.Format("Enter {0}.Execute, DebuggerAttached={1}", this.GetType().Name, this._debuggerIsAttached), this.GetType().Name);

            Guard.IsNotNull(sourceData, "sourceData");

            // Generate the report
            MemoryStreamResult result = this.GenerateReport(sourceData, exportCreationType);
            var resultStatus = result == null ? MemoryStreamResultStatus.Unknown : result.Status;

            _logger.Log(LogType.Trace, this.GetAssemblyName(), string.Format("Exit {0}.Execute, with status {1}", this.GetType().Name, resultStatus), this.GetType().Name);

            return result;
        }

        #region Private Helpers

        private MemoryStreamResult GenerateReport(object sourceData, CreationType exportCreationType)
        {
            var baseUri = GetBaseUri(exportCreationType);

            string reportBaseUri = baseUri + "SampleReport/";
            string reportMetadataFolder = reportBaseUri + "SampleReportMetadata/";
            string reportTemplatesFolder = reportBaseUri + "SampleReportTemplate/";

            // Construct the data part
            var dataParts = new List<ExportDataPart>
            {
                new ExportDataPart
                        {
                            Data = sourceData,
                            PartId = "SampleReportDataPart"
                        }
            };

            // Gather the metadata and template files
            var metadataFiles = new ExportFiles();
            metadataFiles.Metadata.Add("ReportMetaData.xaml", ReadResourceAsString(reportMetadataFolder + "ReportMetaData.xaml"));
            metadataFiles.Templates.Add("ReportMetadataWorkbook.xlsx", ReadResourceAsByteArray(reportMetadataFolder + "ReportMetadataWorkbook.xlsx"));

            var templateFiles = new ExportFiles();
            //templateFiles.Metadata.Add("SharedResources.xaml", ReadResourceAsString(SharedTemplatesFolder + "SharedResources.xaml"));
            templateFiles.Metadata.Add("ReportTemplates.xaml", ReadResourceAsString(reportTemplatesFolder + "ReportTemplates.xaml"));
            templateFiles.Templates.Add("ReportTemplateWorkbook.xlsx", ReadResourceAsByteArray(reportTemplatesFolder + "ReportTemplateWorkbook.xlsx"));

            // Generate the report
            bool isExcelDocumentMetadataBased = exportCreationType == CreationType.ExcelMetadataBased;
            var result = _exportMapService.GenerateReport(dataParts, metadataFiles, templateFiles, isExcelDocumentMetadataBased);
            return result;
        }

        private string GetBaseUri(CreationType creationType)
        {
            // If debugger is attached, this will use the physical disk files (as opposed to assembly resources)
            // so that you don't have to keep re-staring your app while building up your report....!
            string baseUri;
            if (this._debuggerIsAttached)
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                baseUri = "C:/Users/chris_000/Documents/Visual Studio 2013/Projects/Excel Writer/Tests/ExcelWriter Test Harness";
            }
            else
            {
                baseUri = "ExcelWriter.TestHarness";
            }

            baseUri += string.Format("/Maps/{0}/", creationType);
            return baseUri;
        }

        private string ReadResourceAsString(string resourceLocation)
        {
            string result = null;
            Stream stream = this.GetStream(resourceLocation);

            if (stream != null)
            {
                using (var reader = new StreamReader(stream))
                {
                    result = reader.ReadToEnd();
                }
            }

            return result;
        }

        private byte[] ReadResourceAsByteArray(string resourceLocation)
        {
            var result = new byte[0];
            Stream stream = this.GetStream(resourceLocation);

            if (stream != null)
            {
                using (var br = new BinaryReader(stream))
                {
                    result = br.ReadBytes((int)stream.Length);
                }
            }

            return result;
        }

        private Stream GetStream(string resourceLocation)
        {
            Stream stream = null;

            // If debugger is attached, this will use the physical disk files (as opposed to assembly resources)
            // so that you don't have to keep re-staring your app while building up your report....!
            if (this._debuggerIsAttached)
            {
                // Get resource as a file stream 
                using (FileStream fileStream = File.OpenRead(resourceLocation))
                {
                    var memStream = new MemoryStream();
                    memStream.SetLength(fileStream.Length);
                    fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
                    stream = memStream;
                }
            }
            else
            {
                var assembly = Assembly.GetCallingAssembly();
                string resourceName = resourceLocation.Replace("/", ".");
                stream = assembly.GetManifestResourceStream(resourceName);
            }

            return stream;
        }

        #endregion Private Helpers
    }
}
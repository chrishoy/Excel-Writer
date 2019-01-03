namespace ExcelWriter.TestHarness.Maps
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;

    using ExcelWriter;

    /// <summary>
    /// Command which, when executed, creates a position report.
    /// </summary>
    public class CreateExportMetadataBasedReportCommand
    {
        private readonly ILogger _logger;
        private readonly IExportMapService _exportMapService;
        private readonly bool _debuggerIsAttached;

        /// <summary>
        /// Initialises a new instance of the <see cref="CreateSampleReportCommand"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="exportMapService">The export map service.</param>
        public CreateExportMetadataBasedReportCommand(
            ILogger logger,
            IExportMapService exportMapService)
        {
            Guard.IsNotNull(logger, "logger");
            Guard.IsNotNull(exportMapService, "exportMapService");

            _logger = logger;
            _exportMapService = exportMapService;
            _debuggerIsAttached = Debugger.IsAttached;
        }

        public MemoryStreamResult Execute(object sourceData, string partId, string folder)
        {
            _logger.Log(LogType.Trace, this.GetAssemblyName(), string.Format("Enter {0}.Execute, DebuggerAttached={1}", GetType().Name, _debuggerIsAttached), GetType().Name);

            Guard.IsNotNull(sourceData, "sourceData");

            // Generate the report
            MemoryStreamResult result = this.GenerateReport(sourceData, partId, folder);
            var resultStatus = result == null ? MemoryStreamResultStatus.Unknown : result.Status;

            _logger.Log(LogType.Trace, this.GetAssemblyName(), string.Format("Exit {0}.Execute, with status {1}", GetType().Name, resultStatus), GetType().Name);

            return result;
        }

        #region Private Helpers

        private MemoryStreamResult GenerateReport(object sourceData, string partId, string folder)
        {
            var baseUri = GetBaseUri();

            string reportBaseUri = baseUri + folder + "/";
            string reportMetadataFolder = reportBaseUri + "Metadata/";
            string reportTemplatesFolder = reportBaseUri + "Templates/";

            // Construct the data part
            var dataParts = new List<ExportDataPart>
            {
                new ExportDataPart
                        {
                            Data = sourceData,
                            PartId = partId,
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
            const bool isExcelDocumentMetadataBased = false;
            var result = _exportMapService.GenerateReport(dataParts, metadataFiles, templateFiles, isExcelDocumentMetadataBased);
            return result;
        }

        private string GetBaseUri()
        {
            // If debugger is attached, this will use the physical disk files (as opposed to assembly resources)
            // so that you don't have to keep re-staring your app while building up your report....!
            string baseUri;
            if (_debuggerIsAttached)
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                baseUri = "C:/Users/chris_000/Documents/Visual Studio 2013/Projects/Excel Writer/Tests/ExcelWriter Test Harness";
            }
            else
            {
                baseUri = "ExcelWriter.TestHarness";
            }

            baseUri += "/Maps/Samples/";
            return baseUri;
        }

        private string ReadResourceAsString(string resourceLocation)
        {
            string result = null;
            Stream stream = GetStream(resourceLocation);

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
            Stream stream = GetStream(resourceLocation);

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
namespace ExcelWriter.TestHarness
{
    using System.IO;
    using System.Windows;

    using Maps;
    using Maps.Data;
    using ExcelWriter;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly CreateSampleReportCommand _createReportCommand;
        private readonly CreateExportMetadataBasedReportCommand _createExportMetadataBasedReportCommand;
        private IExportMapService _exportMapService;
        private ILogger _logger;

        public MainWindow()
        {
            InitializeComponent();
            _logger = new DebugLogger();
            _exportMapService = new ExportMapService(_logger);
            _createReportCommand = new CreateSampleReportCommand(_logger, _exportMapService);
            _createExportMetadataBasedReportCommand = new CreateExportMetadataBasedReportCommand(_logger, _exportMapService);
        }

        private void ExportMetadataBased_OnClick(object sender, RoutedEventArgs e)
        {
            GenerateAndOpenReport(CreationType.ExportMetadataBased);
        }

        private void ExcelMetadataBased_OnClick(object sender, RoutedEventArgs e)
        {
            GenerateAndOpenReport(CreationType.ExcelMetadataBased);
        }

        private void GenerateStackedBarChartReport_OnClick(object sender, RoutedEventArgs e)
        {
            var data = SampleDataBuilder.BuildStackedBarChartData();
            var result = _createExportMetadataBasedReportCommand.Execute(data, "StackedBarChartDataPart", "StackedBarChart");

            TryWriteAndOpen(result, "StackedBarChart");
        }

        private void GenerateAndOpenReport(CreationType createType)
        {
            var data = SampleDataBuilder.Build();
            var result = this._createReportCommand.Execute(data, createType);

            TryWriteAndOpen(result, createType.ToString());
        }

        private static void TryWriteAndOpen(MemoryStreamResult result, string name)
        {
            switch (result.Status)
            {
                case MemoryStreamResultStatus.Failure:
                    MessageBox.Show(result.ErrorMessage, "Result Status=Failure");
                    break;

                case MemoryStreamResultStatus.Success:
                    // Write to file and open
                    string filePath = string.Format(@"{0}\TestOutput_{1}.xlsx", Directory.GetCurrentDirectory(), name);
                    File.WriteAllBytes(filePath, result.MemoryStream.ToArray());
                    System.Diagnostics.Process.Start(filePath);
                    break;

                default:
                    MessageBox.Show("Unknown Status");
                    break;
            }
        }
    }
}

namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using Drawing = DocumentFormat.OpenXml.Drawing;
    using DrawingCharts = DocumentFormat.OpenXml.Drawing.Charts;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    
    /// <summary>
    /// Encapsulates information about an Excel chart object.
    /// </summary>
    public class ChartModel
    {
        #region Local Fields

        private string chartId;
        private Worksheet worksheet;
        private ChartPart chartPart;
        private DrawingSpreadsheet.TwoCellAnchor anchor;

        private IEnumerable<OpenXmlCompositeElement> chartElements;

        #endregion Local Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartModel"/> class.
        /// </summary>
        /// <param name="chartPart">The <see cref="ChartPart"/> on which the model is based</param>
        /// <param name="anchor">The <see cref="DrawingSpreadsheet.TwoCellAnchor"/> which hosts the chart</param>
        private ChartModel(ChartPart chartPart, DrawingSpreadsheet.TwoCellAnchor anchor)
        {
            this.chartPart = chartPart;
            this.anchor = anchor;
        }

        #endregion

        #region Pubic Properties

        /// <summary>
        /// Gets a value indicating whether the model is valid with the context of the worksheet.
        /// </summary>
        public bool IsValid { get; private set; }

        /// <summary>
        /// Gets the chart Id
        /// </summary>
        public string ChartId
        {
            get { return this.chartId; }
            private set { this.chartId = value; }
        }

        /// <summary>
        /// Gets a reference to the <see cref="ChartPart"/> which this <see cref="ChartModel"/> represents.
        /// </summary>
        public ChartPart ChartPart
        {
            get { return this.chartPart; }
        }

        /// <summary>
        /// Gets the <see cref="Worksheet"/> on which the chart resides
        /// </summary>
        public Worksheet Worksheet
        {
            get { return this.worksheet;  }
            private set { this.worksheet = value; }
        }

        /// <summary>
        /// Gets a list of <see cref="OpenXmlCompositeElement"/>s representing the series on a specified OpenXML chart element.
        /// </summary>
        /// <param name="chartElement">The OpenXML chart element</param>
        /// <returns>A list of <see cref="OpenXmlCompositeElement"/>s representing the series</returns>
        public IEnumerable<OpenXmlCompositeElement> GetSeriesElements(OpenXmlCompositeElement chartElement)
        {
            if (chartElements == null) throw new ArgumentNullException("chartElement");

            List<OpenXmlCompositeElement> seriesElements = new List<OpenXmlCompositeElement>();

            // Add all descendent series in supplied element
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.AreaChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.BarChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.BubbleChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.LineChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.PieChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.RadarChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.ScatterChartSeries>());
            seriesElements.AddRange((IEnumerable<OpenXmlCompositeElement>)chartElement.Descendants<DrawingCharts.SurfaceChartSeries>());

            return seriesElements;
        }

        /// <summary>
        /// Gets a list of ALL <see cref="OpenXmlCompositeElement"/>s representing the series on ALL OpenXML chart elements represented by this model.
        /// </summary>
        /// <returns>A list of ALL <see cref="OpenXmlCompositeElement"/>s representing the series on ALL charts that this model represents</returns>
        public IEnumerable<OpenXmlCompositeElement> GetAllSeriesElements()
        {
            var allSeriesElements = new List<OpenXmlCompositeElement>();

            foreach (var chartElement in this.ChartElements)
            {
                allSeriesElements.AddRange(this.GetSeriesElements(chartElement));
            }

            return allSeriesElements;
        }

        /// <summary>
        /// Gets a list of <see cref="OpenXmlCompositeElement"/>s which represents the charts in this model.
        /// </summary>
        public IEnumerable<OpenXmlCompositeElement> ChartElements
        {
            get { return this.chartElements; }
            private set { this.chartElements = value; }
        }

        #endregion Public Properties

        #region Static Public Methods

        /// <summary>
        /// Returns an instance of a <see cref="ChartModel"/> for the first chart with a specified chart id in a worksheet.
        /// </summary>
        /// <param name="ws">The <see cref="Worksheet"/> in which the chart resides</param>
        /// <param name="id">The chart id</param>
        /// <returns>The <see cref="ChartModel"/> that represents the chart</returns>
        public static ChartModel GetChartModel(Worksheet ws, string id)
        {
            Guard.IsNotNull(ws, "ws");
            Guard.IsNotNullOrEmpty(id, "id");

            ChartModel chartModel = null;
            ChartPart chartPart = GetChartPart(id, ws.WorksheetPart);

            if (chartPart != null)
            {
                // Get the Anchor that host the Graphic
                DrawingSpreadsheet.TwoCellAnchor anchor = GetHostingTwoCellAnchor(chartPart);

                // Get information about the chart
                IEnumerable<OpenXmlCompositeElement> chartElements = GetChartElements(chartPart);

                chartModel = new ChartModel(chartPart, anchor)
                {
                    Worksheet = ws,
                    ChartId = id,
                    ChartElements = chartElements,
                    IsValid = true,
                };
            }

            return chartModel;
        }

        #endregion Static Public Methods

        #region Public Methods

        /// <summary>
        /// Gets the gets the id of a chart in an Excel worksheet for a<see cref="ChartPart"/>.
        /// </summary>
        /// <param name="chartPart">The <see cref="ChartPart"/></param>
        /// <returns>The name/id of the chart in the Excel worksheet.</returns>
        public static string GetIdOfChartPart(ChartPart chartPart)
        {
            DrawingSpreadsheet.TwoCellAnchor twoCellAnchor = GetHostingTwoCellAnchor(chartPart);
            return GetHostedPartName(twoCellAnchor);
        }

        /// <summary>
        /// Creates a deep copy of this <see cref="ChartModel"/> and associated chart in the worksheet.
        /// </summary>
        /// <returns>The <see cref="ChartModel"/> that represents the chart</returns>
        public ChartModel Clone()
        {
            return this.Clone(this.worksheet);
        }

        /// <summary>
        /// Creates a deep copy of this <see cref="ChartModel"/> and associated chart in the worksheet.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet into which the clone will be placed. If null, the cloned <see cref="ChartModel"/> will be based on the original <see cref="Worksheet"/>/></param>
        /// <returns>The <see cref="ChartModel"/> that represents the chart</returns>
        public ChartModel Clone(Worksheet targetWorksheet)
        {
            // If no target worksheet is supplied, clone in situ (ie. on the current worksheet)
            Worksheet cloneToWorksheet = targetWorksheet == null ? this.worksheet : targetWorksheet;

            // Name of the source and target worksheet (for debugging)
            string sourceWorksheetName = this.worksheet.WorksheetPart.GetSheetName();
            string targetWorksheetName = cloneToWorksheet.WorksheetPart.GetSheetName();

            System.Diagnostics.Debug.Print("ChartModel - Cloning chart on worksheet '{0}' into '{1}'", sourceWorksheetName, targetWorksheetName);

            // Create a DrawingPart in the target worksheet if it does not already exist
            if (cloneToWorksheet.WorksheetPart.DrawingsPart == null)
            {
                var drawingsPart = cloneToWorksheet.WorksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new DrawingSpreadsheet.WorksheetDrawing();

                // if a drawings part is being created then we need to add a Drawing to the end of the targetworksheet
                DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                {
                    Id = cloneToWorksheet.WorksheetPart.GetIdOfPart(cloneToWorksheet.WorksheetPart.DrawingsPart)
                };

                cloneToWorksheet.Append(drawing);
            }

            // Take copy elements
            ChartPart chartPart2 = cloneToWorksheet.WorksheetPart.DrawingsPart.AddNewPart<ChartPart>();
            chartPart2.FeedData(this.chartPart.GetStream());

            // Clone the anchor for the template chart to get a new chart anchor
            DrawingSpreadsheet.TwoCellAnchor anchor2 = (DrawingSpreadsheet.TwoCellAnchor)this.anchor.CloneNode(true);

            // Insert the cloned anchor into the worksheet drawing of the DrawingsPart.
            cloneToWorksheet.WorksheetPart.DrawingsPart.WorksheetDrawing.Append(anchor2);

            // Update the ChartReference in the Anchor 2 (TwoCellAnchor -> GraphicFrame -> Graphic -> GraphicData -> ChartReference)
            DrawingCharts.ChartReference chartReference2 = anchor2.Descendants<DrawingCharts.ChartReference>().FirstOrDefault();
            chartReference2.Id = cloneToWorksheet.WorksheetPart.DrawingsPart.GetIdOfPart(chartPart2);

            // Get information about the cloned chart
            IEnumerable<OpenXmlCompositeElement> chartElements = GetChartElements(chartPart2);

            // Wrap and return as a model
            ChartModel chartModel = new ChartModel(chartPart2, anchor2)
            {
                Worksheet = cloneToWorksheet,
                ChartId = this.ChartId,
                ChartElements = chartElements,
                IsValid = true,
            };

            return chartModel;
        }

        /// <summary>
        /// Moves the chart into position within its worksheet.
        /// </summary>
        /// <param name="fromRowIndex">The worksheet row where the chart will start</param>
        /// <param name="fromColumnIndex">The worksheet column where the chart will start</param>
        /// <param name="toRowIndex">The worksheet row when the chart will end</param>
        /// <param name="toColumnIndex">The worksheet column where the chart will end</param>
        public void Move(uint fromRowIndex, uint fromColumnIndex, uint toRowIndex, uint toColumnIndex)
        {
            // From Marker
            this.anchor.FromMarker.RowId.Text = fromRowIndex.ToString();
            this.anchor.FromMarker.RowOffset.Text = "0";
            this.anchor.FromMarker.ColumnId.Text = fromColumnIndex.ToString();
            this.anchor.FromMarker.ColumnOffset.Text = "0";

            // To Marker
            this.anchor.ToMarker.RowId.Text = toRowIndex.ToString();
            this.anchor.ToMarker.RowOffset.Text = "0";
            this.anchor.ToMarker.ColumnId.Text = toColumnIndex.ToString();
            this.anchor.ToMarker.ColumnOffset.Text = "0";
        }

        /// <summary>
        /// Returns the series at the specified zero-based index position in the supplied OpenXML chart element.
        /// </summary>
        /// <typeparam name="SeriesType">The <see cref="OpenXmlCompositeElement"/> derrived type of the chart series</typeparam>
        /// <param name="idx">The 0-based index</param>
        /// <returns>A <see cref="SeriesType"/></returns>
        public SeriesType GetSeries<SeriesType>(OpenXmlCompositeElement chartElement, int idx) where SeriesType : OpenXmlCompositeElement
        {
            IEnumerable<OpenXmlCompositeElement> seriesElements = this.GetSeriesElements(chartElement);

            return (SeriesType)seriesElements.ElementAt(idx);
        }

        /// <summary>
        /// Removes all references to the chart from the worksheet.<br/>
        /// This will invalidate this <see cref="ChartModel"/>, i.e. errors will be raised if an attempt is made to use the invalid model.
        /// </summary>
        public void RemoveChart()
        {
            try
            {
                this.worksheet.WorksheetPart.DrawingsPart.DeletePart(this.chartPart);

                IEnumerable<ChartPart> chartParts = this.worksheet.WorksheetPart.DrawingsPart.GetPartsOfType<ChartPart>();

                // Remove the Anchor from the WorksheetDrawing
                this.anchor.Remove();
                this.anchor = null;

                this.chartPart = null;
                this.worksheet = null;

                this.IsValid = false;
            }
            catch (System.InvalidOperationException)
            {
                // Do nothing, the part had already beed destroyed, probably by another model based on the same chart
            }
        }

        /// <summary>
        /// Updates the title text on the chart currently modelled.<br/>
        /// NB! There are still issues using this technique, eg. title orientation doesn't work if your source chart has title rotated.<br/>
        /// You may wish to provide your own chart title by placing a cell and chart in a vertically orientated stach panel.
        /// </summary>
        /// <param name="title">The text to be set as the title</param>
        /// <returns>true if the title was set, otherwise false</returns>
        public bool SetTitle(string title)
        {
            DrawingCharts.Title chartTitle = this.chartPart.ChartSpace.Descendants<DrawingCharts.Title>().FirstOrDefault();

            if (chartTitle != null)
            {
                // If there are any text properties or ChartText, then remove
                DrawingCharts.TextProperties tp = chartTitle.Descendants<DrawingCharts.TextProperties>().FirstOrDefault();
                if (tp != null)
                {
                    tp.Remove();
                }
                else
                {
                    DrawingCharts.ChartText ct = chartTitle.Descendants<DrawingCharts.ChartText>().FirstOrDefault();
                    {
                        if(ct!= null)
                            ct.Remove();
                    }
                }

                // Insert ChartText
                DrawingCharts.ChartText chartText = CreateChartText(title);
                chartTitle.InsertAt(chartText, 0);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Creates ChartText which can be used to set a title on a chart.
        /// </summary>
        /// <param name="title">The text for the title</param>
        /// <returns>A <see cref="DrawingCharts.ChartText"/></returns>
        private static DrawingCharts.ChartText CreateChartText(string title)
        {
            var chartText = new DrawingCharts.ChartText();

            var richText = new DrawingCharts.RichText();
            var bodyProperties = new Drawing.BodyProperties();
            var listStyle = new Drawing.ListStyle();

            var paragraph = new Drawing.Paragraph();

            var paragraphProperties = new Drawing.ParagraphProperties();
            var defaultRunProperties = new Drawing.DefaultRunProperties();

            paragraphProperties.Append(defaultRunProperties);

            var run = new Drawing.Run();
            var runProperties = new Drawing.RunProperties() { Language = "en-US" };
            var text = new Drawing.Text(title);

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            richText.Append(bodyProperties);
            richText.Append(listStyle);
            richText.Append(paragraph);

            chartText.Append(richText);
            return chartText;
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Gets the Name of a visual graphic element which is hosted in a <see cref="DrawingSpreadsheet.TwoCellAnchor"/>
        /// </summary>
        /// <param name="twoCellAnchor">The <see cref="DrawingSpreadsheet.TwoCellAnchor"/></param>
        /// <returns>The name of the hosted element, or null if not found.</returns>
        private static string GetHostedPartName(DrawingSpreadsheet.TwoCellAnchor twoCellAnchor)
        {
            // This horrendous block of code is there to find the id of an element in a GraphicFrame,
            // which hosts a GraphicElement, such as a ChartPart or a shate - There must be a simpler way...!
            string name = null;
            var graphicFrame = twoCellAnchor.Descendants<DrawingSpreadsheet.GraphicFrame>().FirstOrDefault();

            if (graphicFrame != null &&
                graphicFrame.NonVisualGraphicFrameProperties != null &&
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties != null &&
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.HasValue)
            {
                name = graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.Value;
            }

            return name;
        }

        /// <summary>
        /// Gets the <see cref="DrawingSpreadsheet.TwoCellAnchor"/> that is used to host the
        /// <see cref="DrawingSpreadsheet.GraphicFrame"/> in which the <see cref="ChartPart"/> resides.
        /// </summary>
        /// <param name="chartPart">The <see cref="ChartPart"/></param>
        /// <returns>The hosting <see cref="DrawingSpreadsheet.TwoCellAnchor"/></returns>
        private static DrawingSpreadsheet.TwoCellAnchor GetHostingTwoCellAnchor(ChartPart chartPart)
        {
            // Get the chart reference id for the chartPart
            string chartRefId = GetChartReferenceId(chartPart);

            // DrawingPart is parent of ChartPart, and WorksheetPart parent of DrawingParts
            var drawingsPart = (DrawingsPart)chartPart.GetParentParts().FirstOrDefault();

            // Get chart reference which matches reference id of supplied chart part.
            DrawingCharts.ChartReference chartRef = drawingsPart.WorksheetDrawing.Descendants<DrawingCharts.ChartReference>().FirstOrDefault(cr => cr.Id == chartRefId);

            // Work bak up to, and return, TwoCellAnchor
            DrawingSpreadsheet.TwoCellAnchor anchor = null;
            if (chartRef != null)
            {
                anchor = chartRef.Ancestors<DrawingSpreadsheet.TwoCellAnchor>().FirstOrDefault();
            }

            return anchor;
        }

        /// <summary>
        /// Gets the id of the chart reference for the <see cref="ChartPart"/>
        /// </summary>
        /// <param name="chartPart">The <see cref="ChartPart"/></param>
        /// <returns>The chart reference id</returns>
        private static string GetChartReferenceId(ChartPart chartPart)
        {
            // DrawingPart is parent of ChartPart, and WorksheetPart parent of DrawingParts
            var drawingsPart = chartPart.GetParentParts().FirstOrDefault();
            return drawingsPart.GetIdOfPart(chartPart);
        }

        /// <summary>
        /// Gets all of the OpenXML chart elements within the <see cref="ChartPart"/>
        /// </summary>
        /// <param name="chartPart">The <see cref="ChartPart"/></param>
        /// <returns>A list of <see cref="OpenXmlCompositeElement"/> objects representing the charts in the <see cref="ChartPart"/></returns>
        private static IEnumerable<OpenXmlCompositeElement> GetChartElements(ChartPart chartPart)
        {
            List<OpenXmlCompositeElement> charts = new List<OpenXmlCompositeElement>();

            if (chartPart.ChartSpace != null)
            {
                // Get the charts
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.LineChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.PieChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.Pie3DChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.AreaChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.BarChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.Bar3DChart>());
                charts.AddRange((IEnumerable<OpenXmlCompositeElement>)chartPart.ChartSpace.Descendants<DrawingCharts.ScatterChart>());
            }

            return charts;
        }

        /// <summary>
        /// Gets the identified <see cref="ChartPart"/> on a <see cref="WorksheetPart"/>
        /// </summary>
        /// <param name="id">The id of the chart</param>
        /// <param name="wp">The <see cref="WorksheetPart"/></param>
        /// <returns>The identified <see cref="ChartPart"/></returns>
        private static ChartPart GetChartPart(string id, WorksheetPart wp)
        {
            DrawingSpreadsheet.GraphicFrame sourceFrame = GetHostingGraphicFrame(id, wp);

            // we have the graphics frame with data, so now pull out the chart part
            if (sourceFrame != null && sourceFrame.Graphic != null && sourceFrame.Graphic.GraphicData != null)
            {
                var sourceChartRef = sourceFrame.Graphic.GraphicData.Descendants<DrawingCharts.ChartReference>().FirstOrDefault();
                if (sourceChartRef != null && sourceChartRef.Id.HasValue)
                {
                    return (ChartPart)wp.DrawingsPart.GetPartById(sourceChartRef.Id.Value);
                }
            }

            return null;
        }

        private static DrawingSpreadsheet.GraphicFrame GetHostingGraphicFrame(string id, WorksheetPart wp)
        {
            DrawingSpreadsheet.GraphicFrame sourceFrame = null;
            if (wp.DrawingsPart != null)
            {
                // we need to pull out the graphic frame that matches the supplied name
                foreach (var gf in wp.DrawingsPart.WorksheetDrawing.Descendants<DrawingSpreadsheet.GraphicFrame>())
                {
                    // need to check it has the various properties
                    if (gf.NonVisualGraphicFrameProperties != null &&
                        gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties != null &&
                        gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.HasValue)
                    {
                        // and then try and match
                        if (id.CompareTo(gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.Value) == 0)
                        {
                            sourceFrame = gf;
                            break;
                        }
                    }
                }
            }

            return sourceFrame;
        }

        #endregion
    }
}

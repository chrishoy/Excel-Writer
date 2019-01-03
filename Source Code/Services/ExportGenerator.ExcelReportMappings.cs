namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Media;
    using System.Windows.Documents;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;

    using DrawingCharts = DocumentFormat.OpenXml.Drawing.Charts;
    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

    using OpenXml.Excel;
    using OpenXml.Excel.Model;

    /// <summary>
    /// 
    /// </summary>
    public sealed partial class ExportGenerator
    {


        #region Dynamic chart and mapping processing

        /// <summary>
        /// Processes the dynamic chart.
        /// </summary>
        /// <param name="dataPart">The data part.</param>
        /// <param name="tableDataSheetName">Name of the table data sheet.</param>
        /// <param name="tableData">The table data.</param>
        /// <param name="presentationWSPart">The presentation ws part.</param>
        /// <returns></returns>
        private static bool ProcessDynamicChart(IDataPart dataPart, string tableDataSheetName, TableData tableData, WorksheetPart presentationWSPart)
        {
            // TODO: Check this - At the moment, only one chart is supported
            //   Ie. Export the dataPart to a worksheet, which is married with a presentation worksheet,
            //       which contains a single chart (the chart that is cloned).
            if (dataPart is IPreparable) // Don't think it needs to be preparable anymore.
            {
                var chartPart = presentationWSPart.DrawingsPart.ChartParts.FirstOrDefault();
                if (chartPart != null)
                {                    string id = ChartModel.GetIdOfChartPart(chartPart);

                    // Get the chart we wish to update using a ChartModel, and the last series in the chart, so that we can clone the series.
                    ChartModel chartModel = ChartModel.GetChartModel(presentationWSPart.Worksheet, id);

                    if (chartModel.ChartElements == null)
                    {
                        return false;
                    }

                    int seriesCount = 0;
                    foreach (OpenXmlCompositeElement chartElement in chartModel.ChartElements)
                    {
                        seriesCount += chartModel.GetSeriesElements(chartElement).Count();
                    }

                    if (seriesCount == 0)
                    {
                        return false;
                    }
                    else
                    {
                        return ProcessDynamicChart(tableDataSheetName, tableData, chartModel);
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Processes the dynamic chart.
        /// </summary>
        /// <param name="tableDataSheetName">Name of the table data sheet.</param>
        /// <param name="tableData">The table data.</param>
        /// <param name="chartModel">The chart model.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">chartModel</exception>
        private static bool ProcessDynamicChart(string tableDataSheetName, TableData tableData, ChartModel chartModel)
        {
            if (chartModel == null) throw new ArgumentNullException("chartModel");

            TempDiagnostics.Output(string.Format("Processing chart in sheet '{0}'", tableDataSheetName));

            var seriesFactory = new SeriesFactory(chartModel);

            if (seriesFactory.SourceSeriesCount > 0)
            {
                if (tableData.TreatRowAsSeries == false)
                {
                    // A COLUMN of data represents a series...
                    // Find the Range that representes the first Axis Data for all of the series...
                    CompositeRangeReference category1AxisRange = DetermineCategory1AxisRange(tableData, tableData.TreatRowAsSeries, tableDataSheetName);

                    foreach (TableColumn column in tableData.Columns)
                    {
                        // Skip if we are on the category column
                        if (column.DataRegion.ExcelColumnStart == category1AxisRange.MinColumnIndex)
                        {
                            continue;
                        }

                        // Get information about the series that will be based on the column
                        var seriesInfo = new ChartSeriesInfo(column);

                        // We need a column that has not been actively excluded
                        if (seriesInfo.BaseOnChartSeriesIndex >= 0 && !seriesInfo.SuppressSeries)
                        {
                            // Get the template series or get a copy if already used
                            OpenXmlCompositeElement clonedSeries = seriesFactory.GetOrCloneSourceSeries(seriesInfo.BaseOnChartSeriesIndex);

                            //TODO: Allow ability to assign colour palettes (as opposed to colours) to dynamically generated series
                            Color? seriesColour = seriesInfo.SeriesColour;
                            if (!seriesColour.HasValue)
                            {
                                int useCount = seriesFactory.GetUseCount(seriesInfo.BaseOnChartSeriesIndex);
                                seriesColour = new Color?((ColourPalette.GetColour(ColourPaletteType.GamTechnicalChartPalette, useCount - 1)));
                            }

                            var seriesTextRange = new CompositeRangeReference(new RangeReference(tableDataSheetName,
                                                                        (uint)column.DataRegion.ExcelRowStart,
                                                                        (uint)column.DataRegion.ExcelColumnStart));



                            // Series data is first column only
                            var seriesValuesRange = new CompositeRangeReference(new RangeReference(tableDataSheetName,
                                                                        (uint)column.DataRegion.ExcelRowStart + 1,
                                                                        (uint)column.DataRegion.ExcelColumnStart,
                                                                        (uint)column.DataRegion.ExcelRowEnd,
                                                                        (uint)column.DataRegion.ExcelColumnStart));

                            // Determine data ranges to be used within chart series.
                            ChartDataRangeInfo dataRangeInfo = new ChartDataRangeInfo
                            {
                                SeriesTextRange = seriesTextRange,
                                CategoryAxisDataRange = category1AxisRange,
                                SeriesValuesRange = seriesValuesRange,
                            };

                            // update all formula on the series
                            UpdateChartSeriesDataReferences(clonedSeries, dataRangeInfo, new SolidColorBrush(seriesColour.Value));
                        }
                    }
                }
                else
                {
                    // A ROW of data represents a series...
                    // Find the Range that representes the first Axis Data for all of the series...
                    CompositeRangeReference category1AxisRange = DetermineCategory1AxisRange(tableData, tableData.TreatRowAsSeries, tableDataSheetName);
                    TableColumn seriesTextColumn = tableData.Columns[FindNonExcludedColumnIndex(tableData, 1)];

                    // Build a list of columns to include in the series
                    var seriesTableColumns = new List<TableColumn>();

                    int seriesValuesColumnIndex = FindNonExcludedColumnIndex(tableData, 2);
                    TableColumn firstSeriesValuesColumn = tableData.Columns[seriesValuesColumnIndex];
                    seriesTableColumns.Add(firstSeriesValuesColumn);

                    // Count up to last column, excluding those marked for exclusion
                    while (seriesValuesColumnIndex < (tableData.Columns.Count - 1))
                    {
                        seriesValuesColumnIndex++;
                        TableColumn tableColumn = tableData.Columns[seriesValuesColumnIndex];

                        ChartExcludeOption excludeOption = tableColumn.ChartOptions.GetOptionOrDefault<ChartExcludeOption>();
                        if (excludeOption == null || excludeOption.Exclude == false)
                        {
                            seriesTableColumns.Add(tableColumn);
                        }
                    }

                    foreach (TableDataRowInfo rowInfo in tableData.RowData)
                    {
                        // Extract chart series related properties from row data
                        var seriesInfo = new ChartSeriesInfo(rowInfo.RowData);

                        // Determine where data has been written into Excel
                        uint rowIndex = (uint)tableData.MapContainer.ExcelRowStart + rowInfo.TableRowIndex - 1;

                        // Skip if we are on the category column
                        if (rowIndex == category1AxisRange.MinRowIndex)
                        {
                            continue;
                        }

                        // We need a column that has not been actively excluded
                        if (seriesInfo.BaseOnChartSeriesIndex >= 0 && !seriesInfo.SuppressSeries)
                        {
                            // Get the template series or get a copy if already used
                            OpenXmlCompositeElement clonedSeries = seriesFactory.GetOrCloneSourceSeries(seriesInfo.BaseOnChartSeriesIndex);

                            //TODO: Allow ability to assign colour palettes (as opposed to colours) to dynamically generated series
                            Color? seriesColour = seriesInfo.SeriesColour;
                            if (!seriesColour.HasValue)
                            {
                                int useCount = seriesFactory.GetUseCount(seriesInfo.BaseOnChartSeriesIndex);
                                seriesColour = new Color?((ColourPalette.GetColour(ColourPaletteType.GamTechnicalChartPalette, useCount - 1)));
                            }

                            // SeriesText is the series heading used in legends
                            var seriesTextRange = new CompositeRangeReference(new RangeReference(tableDataSheetName,
                                                                        rowIndex,
                                                                        (uint)seriesTextColumn.DataRegion.ExcelColumnStart));

                            // From column after category to last column in table.
                            var seriesValuesRange = new CompositeRangeReference();
                            foreach (TableColumn tableColumn in seriesTableColumns)
                            {
                                seriesValuesRange.Update(tableDataSheetName,
                                                            rowIndex,
                                                            (uint)tableColumn.DataRegion.ExcelColumnStart,
                                                            rowIndex,
                                                            (uint)tableColumn.DataRegion.ExcelColumnStart);
                            }

                            // Determine data ranges to be used within chart series.
                            ChartDataRangeInfo dataRangeInfo = new ChartDataRangeInfo
                            {
                                SeriesTextRange = seriesTextRange,
                                CategoryAxisDataRange = category1AxisRange,
                                SeriesValuesRange = seriesValuesRange,
                            };

                            // update all formula on the series ( I know we're constantly swapping between brush and color here.... Address later)
                            UpdateChartSeriesDataReferences(clonedSeries, dataRangeInfo, new SolidColorBrush(seriesColour.Value));
                        }
                    }
                }

                // Remove all un-used template series and set the order of remaining.
                uint seriesIndex = 0;
                for (int idx = 0; idx < seriesFactory.SourceSeriesCount; idx++)
                {
                    OpenXmlCompositeElement templateSeries = seriesFactory.GetSourceSeriesElement(idx);
                    
                    // Value is the use-count, if zero then the template needs removing.
                    if (seriesFactory.GetUseCount(idx) == 0)
                    {
                        // Remove the template series.
                        templateSeries.Remove();
                    }
                    else
                    {
                        // Set template series index and order to initial value, then set clones
                        var templateIndex = templateSeries.Descendants<DrawingCharts.Index>().FirstOrDefault();
                        if (templateIndex != null) templateIndex.Val = seriesIndex;

                        var templateOrder = templateSeries.Descendants<DrawingCharts.Order>().FirstOrDefault();
                        if (templateOrder != null) templateOrder.Val = seriesIndex;

                        seriesIndex++;

                        OpenXmlElement lastElement = templateSeries;

                        foreach (var clonedElement in seriesFactory.GetClonedSeriesElements(idx))
                        {
                            var index = clonedElement.Descendants<DrawingCharts.Index>().FirstOrDefault();
                            if (index != null) index.Val = seriesIndex;

                            var order = clonedElement.Descendants<DrawingCharts.Order>().FirstOrDefault();
                            if (order != null) order.Val = seriesIndex;

                            lastElement.InsertAfterSelf<OpenXmlElement>(clonedElement);
                            lastElement = clonedElement;
                            seriesIndex++;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Gets the type of a series so that we know how to treat it when cloned.
        /// (ie. how to update formulae and set colours, such as trend-line colours)
        /// </summary>
        /// <param name="chartSeries">The chart series.</param>
        /// <returns></returns>
        private static SeriesType GetSeriesType(OpenXmlCompositeElement chartSeries)
        {
            if (chartSeries is DrawingCharts.BarChartSeries)
            {
                return SeriesType.Line;
            }

            if (chartSeries is DrawingCharts.AreaChartSeries)
            {
                return SeriesType.Line;
            }

            if (chartSeries is DrawingCharts.BubbleChartSeries)
            {
                return SeriesType.Line;
            }

            if (chartSeries is DrawingCharts.LineChartSeries)
            {
                return SeriesType.Line;
            }

            if (chartSeries is DrawingCharts.PieChartSeries)
            {
                return SeriesType.Pie;
            }

            if (chartSeries is DrawingCharts.RadarChartSeries)
            {
                return SeriesType.Line;
            }

            if (chartSeries is DrawingCharts.ScatterChartSeries)
            {
                return SeriesType.Scatter;
            }

            if (chartSeries is DrawingCharts.SurfaceChartSeries)
            {
                return SeriesType.Line;
            }

            return SeriesType.Unrecognised;
        }

        /// <summary>
        /// Updates the supplied series in a chart with range related information.
        /// </summary>
        /// <param name="chartSeries">The OpenXml representation of the chart which is to be updated</param>
        /// <param name="dataRangeInfo">Information relating to the category, axis and series data ranges.</param>
        /// <param name="seriesBrush">The brush which will be used when creating the series in the chart.</param>
        private static void UpdateChartSeriesDataReferences(OpenXmlCompositeElement chartSeries, ChartDataRangeInfo dataRangeInfo, Brush seriesBrush)
        {

            SeriesType seriesType = GetSeriesType(chartSeries);

            switch (seriesType)
            {
                case SeriesType.Line:
                    {
                        // Update series to match Excel document (leave series index/order 0 as will be set later)
                        chartSeries.UpdateCategoryValueChartSeries((uint)0, // Series Index
                                                                    dataRangeInfo.SeriesTextRange,
                                                                    dataRangeInfo.CategoryAxisDataRange,
                                                                    dataRangeInfo.SeriesValuesRange);
                        chartSeries.UpdateLineBrush(seriesBrush);
                        break;
                    }
                case SeriesType.Pie:
                    {
                        // Update series to match Excel document (leave series index/order 0 as will be set later)
                        chartSeries.UpdateCategoryValueChartSeries((uint)0, // Series Index
                                                                    dataRangeInfo.SeriesTextRange,
                                                                    dataRangeInfo.CategoryAxisDataRange,
                                                                    dataRangeInfo.SeriesValuesRange);
                        //TODO: Update based on the value, not the series....
                        // chartSeries.UpdateLineBrush(seriesBrush);
                        break;
                    }
                case SeriesType.Scatter:
                    {
                        chartSeries.UpdateSeriesMarkerBrush(seriesBrush);

                        // if the scatter has a trendline then update its brush to match
                        var trendline = chartSeries.Descendants<DrawingCharts.Trendline>().FirstOrDefault();
                        if (trendline != null)
                        {
                            trendline.UpdateLineBrush(seriesBrush);
                        }

                        ((DrawingCharts.ScatterChartSeries)chartSeries).UpdateXYValueChartSeries((uint)0, // Series Index
                                                                                    dataRangeInfo.SeriesTextRange, // Category Axis column Index
                                                                                    dataRangeInfo.CategoryAxisDataRange, // Data Column Index
                                                                                    dataRangeInfo.SeriesValuesRange);
                        break;
                    }
            }
        }

        /// <summary>
        /// Finds the nth available column in a <see cref="TableData" /><br />
        /// This is simply the nth non-excluded column in the TableData
        /// </summary>
        /// <param name="tableData">The Table to be searched.</param>
        /// <param name="which">The which.</param>
        /// <returns>
        /// A <see cref="TableColumn" />
        /// </returns>
        /// <exception cref="InvalidOperationException"></exception>
        private static int FindNonExcludedColumnIndex(TableData tableData, int which)
        {
            // Find the index of the first column which is not marked as Exluded
            int foundColumnCount = 0;
            int foundColumnIdx = -1;

            for (int idx = 0; idx < tableData.Columns.Count; idx++)
            {
                TableColumn col = tableData.Columns[idx];
                ChartExcludeOption colExcludedOption = col.ChartOptions.GetOptionOrDefault<ChartExcludeOption>();

                // No exclude option set, or exclude option set to false
                if (colExcludedOption == null || colExcludedOption.Exclude == false)
                {
                    foundColumnCount++;
                    if (foundColumnCount == which)
                    {
                        foundColumnIdx = idx;
                        break;
                    }
                }
            }

            if (foundColumnIdx == -1)
            {
                throw new InvalidOperationException(string.Format("The first {0} (read 1 as first, 2 as second etc) non-excluded column does not exist in the TableData.", which));
            }

            return foundColumnIdx;
        }

        /// <summary>
        /// Determines the Category (or X1) Axis column within the table by doing the following:
        /// 1. If there are no columns, then returns null
        /// 2. Iterates over each column, if 'IsCategory1Axis == true', then returns that column.<br />
        /// 3. If no column with 'IsCategory1Axis == true', then assumes the first column, hence returns column 0.
        /// </summary>
        /// <param name="tableData">The table to be tested</param>
        /// <param name="treatRowAsSeries">if set to <c>true</c> [treat row as series].</param>
        /// <param name="dataSheetName">Name of the data sheet.</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">There are no non-excluded columns which can be used as the 'Category 1' axis in chart</exception>
        private static CompositeRangeReference DetermineCategory1AxisRange(TableData tableData, bool treatRowAsSeries, string dataSheetName)
        {
            if (treatRowAsSeries)
            {
                int rowIndex = -1;

                // When rows are treated as series, the Category 1 Axis is the row that contains the row headings,
                // unless a row exists which is explicitly marked as being the Category 1 Axis
                foreach (TableDataRowInfo row in tableData.RowData)
                {
                    // Get series related information about the row
                    var rowInfo = new ChartSeriesInfo(row.RowData);

                    if (rowInfo.IsCategory1Axis)
                    {
                        rowIndex = (int)row.TableRowIndex;
                    }
                }

                if (rowIndex == -1)
                {
                    // No row marked to be used as the Category 1 Axis, so we assume the table heading row
                    rowIndex = (int)tableData.MapContainer.ExcelRowStart;
                }

                // Now we've determined which row will be used as the 'Category 1 Axis',
                // we need to find the first non-excluded column in which the values reside.
                int category1AxisColumnIndex = FindNonExcludedColumnIndex(tableData, 2);
                TableColumn firstColumn = tableData.Columns[category1AxisColumnIndex];

                var rangeReference = new CompositeRangeReference
                                            (
                                                dataSheetName,
                                                (uint)rowIndex,
                                                (uint)firstColumn.DataRegion.ExcelColumnStart,
                                                (uint)rowIndex,
                                                (uint)firstColumn.DataRegion.ExcelColumnStart
                                            );

                // Count up to last column excluding those marked excluded
                while (category1AxisColumnIndex < (tableData.Columns.Count - 1))
                {
                    category1AxisColumnIndex++;
                    TableColumn tableColumn = tableData.Columns[category1AxisColumnIndex];

                    ChartExcludeOption excludeOption = tableColumn.ChartOptions.GetOptionOrDefault<ChartExcludeOption>();
                    if (excludeOption == null || excludeOption.Exclude == false)
                    {
                        rangeReference.Update(dataSheetName,
                                              (uint)rowIndex,
                                              (uint)tableColumn.DataRegion.ExcelColumnStart,
                                              (uint)rowIndex,
                                              (uint)tableColumn.DataRegion.ExcelColumnStart);
                    }
                }

                // Again (see above), we are currently ignoring excluded column in the table from this range reference.
                return rangeReference;
            }
            else
            {
                // Axis Data Range is first non-excluded column data, unless a column is explitly marked as being the Category 1 Axis
                TableColumn dataColumn = null;

                if (tableData.Columns.Count == 0) return null;

                TableColumn firstNonExcludedColumn = null;
                foreach (TableColumn column in tableData.Columns)
                {
                    var colInfo = new ChartSeriesInfo(column);
                    if (colInfo.IsCategory1Axis)
                    {
                        dataColumn = column;
                        break;
                    }

                    if (firstNonExcludedColumn == null && !colInfo.SuppressSeries)
                    {
                        firstNonExcludedColumn = column;
                    }
                }

                if (dataColumn == null)
                {
                    dataColumn = firstNonExcludedColumn;
                }

                // Throw exception if no axis column defined.
                if (dataColumn == null)
                {
                    throw new InvalidOperationException("There are no non-excluded columns which can be used as the 'Category 1' axis in chart");
                }

                return new CompositeRangeReference
                           (
                                new RangeReference(dataSheetName,
                                          (uint)dataColumn.DataRegion.ExcelRowStart + 1,
                                          (uint)dataColumn.DataRegion.ExcelColumnStart,
                                          (uint)dataColumn.DataRegion.ExcelRowEnd,
                                          (uint)dataColumn.DataRegion.ExcelColumnEnd)
                           );
            }
        }

        /// <summary>
        /// Using the range mapping on the supplied report part copy a set a rows
        /// from the source worksheet part to the target worksheet part
        /// </summary>
        /// <param name="set">The set.</param>
        /// <param name="mapping">The mapping.</param>
        /// <param name="sourceWorksheetPart">The source worksheet part.</param>
        /// <param name="targetWorksheetPart">The target worksheet part.</param>
        private static void ProcessRangeMapping(ExportTripleSet set, RangeMapping mapping, WorksheetPart sourceWorksheetPart, WorksheetPart targetWorksheetPart)
        {
            // if there's no row count on the mapping then default from the data part's row count
            int rowCount = mapping.RowCount.HasValue ? mapping.RowCount.Value : set.DataPart.RowCount;

            // set up source and target sheet data
            var sourceSheetData = sourceWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();
            var targetSheetData = targetWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();

            // set up source and target merge cells (this needs to be maintained if any merge cells in the source table thats being mapped)
            var sourceMergeCells = sourceWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.MergeCells>();
            var targetMergeCells = targetWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.MergeCells>();

            // set counters
            // taking the source and target row indices from the range mapping object
            int sourceCounter = mapping.SourceStartRowIndex, targetCounter = mapping.TargetStartRowIndex;

            // get source rows for our range
            var sourceRows = from s in sourceSheetData.Elements<OpenXmlSpreadsheet.Row>()
                             where s.RowIndex >= (mapping.SourceStartRowIndex + 1)
                             && s.RowIndex <= (mapping.SourceStartRowIndex + rowCount)
                             select s;

            // get target rows for our range
            var targetRows = from t in targetSheetData.Elements<OpenXmlSpreadsheet.Row>()
                             where t.RowIndex >= (mapping.TargetStartRowIndex + 1)
                             && t.RowIndex <= (mapping.TargetStartRowIndex + rowCount)
                             select t;

            // this could be negative
            int rowOffset = mapping.TargetStartRowIndex - mapping.SourceStartRowIndex;

            //
            foreach (var sourceRow in sourceRows)
            {
                var matchTarget = (from m in targetRows
                                   where m.RowIndex.HasValue
                                   && m.RowIndex.Value == (uint)(sourceRow.RowIndex + rowOffset)
                                   select m).FirstOrDefault();

                // if there's a target row, get rid, the source is going to cloned
                if (matchTarget != null)
                {
                    targetSheetData.RemoveChild<OpenXmlSpreadsheet.Row>(matchTarget);
                }

                // clone the source
                matchTarget = sourceRow.CloneNode(true) as OpenXmlSpreadsheet.Row;
                // at set the cloned targetn row index correctly (this is not zero based)
                matchTarget.RowIndex = new UInt32Value((uint)(sourceRow.RowIndex + rowOffset));

                // foreach cell in this row update any cell references and formula
                foreach (var cell in matchTarget.Descendants<OpenXmlSpreadsheet.Cell>())
                {
                    // there is a cell reference
                    if (cell.CellReference.HasValue)
                    {
                        // replace the index of the source with the target
                        string newCellReference = cell.CellReference.Value.Replace((sourceRow.RowIndex).ToString(), (matchTarget.RowIndex).ToString());

                        // need to move across any merge cells from source to target
                        if (sourceMergeCells != null)
                        {
                            var matchSourceMergeCells = from smc in sourceMergeCells.Descendants<OpenXmlSpreadsheet.MergeCell>()
                                                        where smc.Reference.HasValue
                                                        && smc.Reference.Value.Contains(cell.CellReference.Value)
                                                        select smc;

                            if (targetMergeCells != null)
                            {
                                foreach (var mc in matchSourceMergeCells)
                                {
                                    // make sure there isnt already one
                                    var matchTargetMergeCells = (from tmc in targetMergeCells.Descendants<OpenXmlSpreadsheet.MergeCell>()
                                                                 where tmc.Reference.HasValue
                                                                 && tmc.Reference.Value.Contains(newCellReference)
                                                                 select tmc).FirstOrDefault();

                                    // if not
                                    if (matchTargetMergeCells == null)
                                    {
                                        // clone
                                        var targetMergeCell = mc.CloneNode(true) as OpenXmlSpreadsheet.MergeCell;
                                        // and set the reference 
                                        targetMergeCell.Reference = new StringValue(mc.Reference.Value.Replace((sourceRow.RowIndex).ToString(), (matchTarget.RowIndex).ToString()));
                                        targetMergeCells.Append(targetMergeCell);
                                    }
                                }
                            }
                        }

                        cell.CellReference.Value = newCellReference;

                    }
                    // there is a cell formula
                    if (cell.CellFormula != null && !string.IsNullOrEmpty(cell.CellFormula.Text))
                    {
                        // update using the new sheet name
                        cell.CellFormula.Text = Helpers.UpdateFormula(cell.CellFormula.Text, set.Template.DataTemplateSheet, set.Part.DataSheetName, null);
                    }
                }

                // make sure the new row is the correct sequence
                var prevRow = (from r in targetSheetData.Descendants<OpenXmlSpreadsheet.Row>()
                               where r.RowIndex < matchTarget.RowIndex
                               select r).LastOrDefault();

                // if there is now prev found just append
                if (prevRow == null)
                {
                    targetSheetData.Append(matchTarget);
                }
                else
                {
                    // otherwise insert after
                    targetSheetData.InsertAfter<OpenXmlSpreadsheet.Row>(matchTarget, prevRow);
                }
            }
        }

        /// <summary>
        /// try and find the 1st placeholder in the set with an existing shape
        /// as they get used they get removed
        /// </summary>
        /// <param name="set">The set.</param>
        /// <param name="placeholders">The placeholders.</param>
        /// <param name="document">The document.</param>
        /// <returns></returns>
        private static MappingPlaceholder GeNextPlaceholderFromSet(MappingPlaceholderSet set, Dictionary<string, MappingPlaceholder> placeholders, SpreadsheetDocument document)
        {
            foreach (var item in set.Items)
            {
                if (!placeholders.ContainsKey(item.MappingPlaceholderId))
                {
                    continue;
                }

                var dp = placeholders[item.MappingPlaceholderId];
                var worksheetPart = document.WorkbookPart.Workbook.GetWorksheetPartByName(dp.SheetName);

                DrawingSpreadsheet.Shape matchShape = worksheetPart.GetShapeByName(dp.Id);
                if (matchShape != null)
                {
                    return dp;
                }
            }
            return null;
        }

        /// <summary>
        /// Copies a chart from the source worksheet part to the report worksheet part
        /// The placeholder id in the drawing mapping information is used to
        /// </summary>
        /// <param name="part">The part.</param>
        /// <param name="placeholder">The placeholder.</param>
        /// <param name="mapping">The mapping.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="reportWSPart">The report ws part.</param>
        /// <returns></returns>
        private static ChartPart ProcessDrawingPart(ExportPart part, MappingPlaceholder placeholder, DrawingMapping mapping, WorksheetPart worksheetPart, WorksheetPart reportWSPart)
        {
            // if no placeholder is specified then return now
            if (placeholder == null)
            {
                return null;
            }

            DrawingSpreadsheet.Shape matchShape = reportWSPart.GetShapeByName(placeholder.Id);
            if (matchShape == null)
            {
                return null;
            }

            // save the location of this shape, 
            // this information will be used to position the incoming chart
            Extents extents = matchShape.ShapeProperties.Transform2D.Extents;
            Offset offset = matchShape.ShapeProperties.Transform2D.Offset;

            // get the source chart.... if a source drawing id is specified use it
            // otherwise get the 1st one
            ChartPart sourceChartPart = null;
            if (string.IsNullOrEmpty(mapping.SourceDrawingId))
            {
                // the 1st chart if there is one
                sourceChartPart = worksheetPart.DrawingsPart.GetPartsOfType<ChartPart>().FirstOrDefault();
            }
            else
            {
                DrawingSpreadsheet.GraphicFrame sourceFrame = null;

                // we need to pull out the graphic frame that matches the supplied name
                foreach (var gf in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DrawingSpreadsheet.GraphicFrame>())
                {
                    // need to check it has the various properties
                    if (gf.NonVisualGraphicFrameProperties != null &&
                        gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties != null &&
                        gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.HasValue)
                    {
                        // and then try and match
                        if (mapping.SourceDrawingId.CompareTo(gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.Value) == 0)
                        {
                            sourceFrame = gf;
                            break;
                        }
                    }
                }

                // we have the graphics frame with data, so no pull out the chart part
                if (sourceFrame != null && sourceFrame.Graphic != null && sourceFrame.Graphic.GraphicData != null)
                {
                    var sourceChartRef = sourceFrame.Graphic.GraphicData.Descendants<DrawingCharts.ChartReference>().FirstOrDefault();
                    if (sourceChartRef != null && sourceChartRef.Id.HasValue)
                    {
                        sourceChartPart = (ChartPart)worksheetPart.DrawingsPart.GetPartById(sourceChartRef.Id.Value);
                    }
                }
            }

            // create a new one
            var targetChartPart = reportWSPart.DrawingsPart.AddNewPart<ChartPart>();

            // and feed data from old to new
            targetChartPart.FeedData(sourceChartPart.GetStream());

            // chart in now
            // just need the drawing to host it

            // get the graphic frame from the source anchor
            var sourceAnchor = worksheetPart.DrawingsPart.WorksheetDrawing.GetFirstChild<DrawingSpreadsheet.TwoCellAnchor>();
            var sourceGraphicFrame = sourceAnchor.Descendants<DrawingSpreadsheet.GraphicFrame>().FirstOrDefault();

            // add it to the target anchor (ie. the one with the shape removed)
            var targetGraphicFrame = sourceGraphicFrame.CloneNode(true);

            // positon the new graphic frame after the shape its going to replace
            matchShape.Parent.InsertAfter<OpenXmlElement>(targetGraphicFrame, matchShape);

            // and remove the shape, not needed anymore
            matchShape.Remove();

            // update the extents and offsets that were saved above
            var transform = targetGraphicFrame.Descendants<DrawingSpreadsheet.Transform>().FirstOrDefault();
            if (transform != null)
            {
                transform.Extents.Cx = extents.Cx;
                transform.Extents.Cy = extents.Cy;
                transform.Offset.X = offset.X;
                transform.Offset.Y = offset.Y;
            }

            // ensure that the id of the chart reference in the cloned graphic frame matches that of the new cloned chart
            // if this isnt done then no chart will appeat
            var chartReference = targetGraphicFrame.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
            chartReference.Id = reportWSPart.DrawingsPart.GetIdOfPart(targetChartPart);

            return targetChartPart;
        }



        /// <summary>
        /// Processes the paragraph part.
        /// </summary>
        /// <param name="wkb">The WKB.</param>
        /// <param name="part">The part.</param>
        /// <param name="placeholder">The placeholder.</param>
        /// <param name="mapping">The mapping.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="reportWSPart">The report ws part.</param>
        /// <returns></returns>
        private static DrawingSpreadsheet.Shape ProcessParagraphPart(OpenXmlSpreadsheet.Workbook wkb, ExportPart part, MappingPlaceholder placeholder, ParagraphMapping mapping, WorksheetPart worksheetPart, WorksheetPart reportWSPart)
        {
            string SheetName = "";
            uint RowStart = 0;
            uint RowEnd = 0;
            uint ColStart = 0;
            uint ColEnd = 0;

            // if no placeholder is specified then return now
            if (placeholder == null)
            {
                return null;
            }

            DrawingSpreadsheet.Shape matchShape = worksheetPart.GetShapeByName(placeholder.Id);
            if (matchShape == null)
            {
                return null;
            }

            // save the location of this shape, 
            // this information will be used to position the incoming chart
            Extents extents = matchShape.ShapeProperties.Transform2D.Extents;
            Offset offset = matchShape.ShapeProperties.Transform2D.Offset;

            OpenXmlSpreadsheet.DefinedName rtfXmlDefinedName = wkb.GetDefinedNameByName(string.Format("{0}_{1}", part.DataSheetName, mapping.SourceFieldName));

            wkb.BreakDownDefinedName(rtfXmlDefinedName, ref SheetName, ref RowStart, ref RowEnd, ref ColStart, ref ColEnd);

            WorksheetPart wksp = wkb.GetWorksheetPartByName(SheetName);

            OpenXmlSpreadsheet.SheetData sheetData = wksp.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();

            OpenXmlSpreadsheet.Cell rtfXmlCell = sheetData.GetCell(ColStart, RowStart + 1);

            // Use the cell on the hidden data sheet as source for the XAML reader
            Section RTFSection = XamlSectionDocumentReader(rtfXmlCell.CellValue.InnerText);

            // The paragraph in the cell.inlinestring have a very different class structure to the paragraphs in the shape.textbody
            // So, the paragraph will need to go through a converter to do this.

            DrawingSpreadsheet.Shape targetShape = ConvertParagraph(worksheetPart, RTFSection, matchShape);

            // positon the new graphic frame after the shape its going to replace
            matchShape.Parent.InsertAfter<OpenXmlElement>(targetShape, matchShape);

            matchShape.Remove();

            return targetShape;
        }



        /// <summary>
        /// Creates a 'legend' based on the chart part supplied and positions in the placeholder
        /// Useful for single legend shared across multiple charts
        /// </summary>
        /// <param name="chartPart">The chart part.</param>
        /// <param name="legendPlaceholderId">The legend placeholder identifier.</param>
        private static void ProcessLegend(ChartPart chartPart, string legendPlaceholderId)
        {
            // access the chart 
            //   for each series find the series text..... and the colour
            //   then build the legend
        }

        #endregion
    }
}

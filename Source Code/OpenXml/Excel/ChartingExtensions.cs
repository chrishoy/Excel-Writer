namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using DocumentFormat.OpenXml;
    using System.Windows.Media;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Extension methods for OpenXML Chart manipulation.
    /// </summary>
    public static class ChartingExtensions
    {
        /// <summary>
        /// Updates the sources.
        /// </summary>
        /// <param name="chartPart">The chart part.</param>
        /// <param name="oldSourceName">Old name of the source.</param>
        /// <param name="newSourceName">New name of the source.</param>
        /// <param name="tableRowCount">The table row count.</param>
        public static void UpdateSources(this ChartPart chartPart, string oldSourceName, string newSourceName, int? tableRowCount)
        {
            // try and get a chart
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            // if there is one
            if (chart != null)
            {
                // update any sources on it to ensure formula and ranges are correct
                Helpers.UpdateDataSourcesForChildren(chart, oldSourceName, newSourceName, tableRowCount);
            }
        }

        /// <summary>
        /// Updates a supplied <see cref="SolidFill" /> with a <see cref="SolidColorBrush" /> colour.
        /// </summary>
        /// <param name="solidFill">The <see cref="SolidFill" /></param>
        /// <param name="brush">The <see cref="SolidColorBrush" /> which contains the colour which needs to be set</param>
        private static void UpdateSolidFill(this SolidFill solidFill, SolidColorBrush brush)
        {
            solidFill.RemoveAllChildren();

            var scb = (SolidColorBrush)brush;

            StringBuilder hexString = new StringBuilder();
            hexString.Append(brush.Color.R.ToString("X"));
            hexString.Append(brush.Color.G.ToString("X"));
            hexString.Append(brush.Color.B.ToString("X"));

            RgbColorModelHex hexColour = new RgbColorModelHex()
            {
                Val = hexString.ToString()
            };

            solidFill.Append(hexColour);
        }

        /// <summary>
        /// Updates the colour of a <see cref="BarChartSeries" /> so it matches a <see cref="SolidColorBrush" />
        /// </summary>
        /// <param name="series">The series.</param>
        /// <param name="brush">The brush.</param>
        private static void UpdateSeriesColour(this BarChartSeries series, SolidColorBrush brush)
        {
            var chartShapeProperties = series.Descendants<ChartShapeProperties>().FirstOrDefault();
            if (chartShapeProperties == null)
            {
                chartShapeProperties = new ChartShapeProperties();
                // if there's a series text then insert afterwards
                // the series title, this is the name of the column header
                var seriesText = series.Descendants<SeriesText>().FirstOrDefault();
                if (seriesText == null)
                {
                    series.InsertAt<ChartShapeProperties>(chartShapeProperties, 0);
                }
                else
                {
                    series.InsertAfter<ChartShapeProperties>(chartShapeProperties, seriesText);
                }
            }

            // Determine if the series has a SolidFill
            SolidFill seriesSolidFill = chartShapeProperties.Elements<SolidFill>().FirstOrDefault();
            if (seriesSolidFill == null)
            {
                // No fill, so create one so we can set the colour
                seriesSolidFill = new SolidFill();
                seriesSolidFill.UpdateSolidFill((SolidColorBrush)brush);
                chartShapeProperties.InsertAt(seriesSolidFill, 0);
            }
            else
            {
                seriesSolidFill.UpdateSolidFill((SolidColorBrush)brush);
            }
        }

        /// <summary>
        /// Sets the color of the series using a solidcolor brush
        /// If a null brush is supplied any color is removed so the color will be automatic
        /// </summary>
        /// <param name="line">The line.</param>
        /// <param name="brush">The brush.</param>
        public static void UpdateLineBrush(this OpenXmlCompositeElement line, Brush brush)
        {
            if (line == null)
            {
                return;
            }

            // If we have a BarChart, we really want tp update the SolidFill (not the Outline.SolidFill)
            BarChartSeries barChartSeries = line as BarChartSeries;
            if (barChartSeries != null)
            {
                // For BarCharts, we update the SolidFill
                barChartSeries.UpdateSeriesColour((SolidColorBrush)brush);
            }
            else
            {
                // Update the Outline.SolidFill + set the SolidFill to the same colour
                var chartShapeProperties = line.Descendants<ChartShapeProperties>().FirstOrDefault();

                if (brush == null && !(brush is SolidColorBrush))
                {
                    if (chartShapeProperties != null)
                    {
                        line.RemoveChild<ChartShapeProperties>(chartShapeProperties);
                    }
                    return;
                }

                // the series title, this is the name of the column header
                var seriesText = line.Descendants<SeriesText>().FirstOrDefault();

                if (chartShapeProperties == null)
                {
                    chartShapeProperties = new ChartShapeProperties();
                    // if there's a series text then insert afterwards
                    if (seriesText == null)
                    {
                        line.InsertAt<ChartShapeProperties>(chartShapeProperties, 0);
                    }
                    else
                    {
                        line.InsertAfter<ChartShapeProperties>(chartShapeProperties, seriesText);
                    }
                }

                var outline = chartShapeProperties.Descendants<Outline>().FirstOrDefault();
                if (outline == null)
                {
                    outline = new Outline();
                    chartShapeProperties.InsertAt(outline, 0);
                }

                var outlineSolidFill = outline.Descendants<SolidFill>().FirstOrDefault();
                if (outlineSolidFill == null)
                {
                    outlineSolidFill = new SolidFill();
                    outline.Append(outlineSolidFill);
                }

                // Update the fill with the supplied brush colour
                outlineSolidFill.UpdateSolidFill((SolidColorBrush)brush);

                // Clones the OutlineSolidFill as the SolidFill of the series...
                var solidFill = chartShapeProperties.GetFirstChild<SolidFill>();
                if (solidFill != null)
                {
                    chartShapeProperties.RemoveChild(solidFill);
                }
                chartShapeProperties.InsertAt(outlineSolidFill.CloneNode(true), 0);
            }
        }

        /// <summary>
        /// Sets the color of the series using a solidcolor brush
        /// If a null brush is supplied any color is removed so the color will be automatic
        /// </summary>
        /// <param name="series">The series.</param>
        /// <param name="brush">The brush.</param>
        /// <exception cref="ArgumentNullException">series</exception>
        public static void UpdateSeriesMarkerBrush(this OpenXmlCompositeElement series, Brush brush)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            var marker = series.Descendants<Marker>().FirstOrDefault();
            if (marker == null)
            {
                return;
            }

            var scb = brush as SolidColorBrush;
            if (scb == null)
            {
                return;
            }

            // clear down and start again
            marker.RemoveAllChildren();

            var chartShapeProperties = new ChartShapeProperties();

            SolidFill solidFill = new SolidFill();

            StringBuilder hexString = new StringBuilder();
            hexString.Append(scb.Color.R.ToString("X"));
            hexString.Append(scb.Color.G.ToString("X"));
            hexString.Append(scb.Color.B.ToString("X"));

            RgbColorModelHex hexColour = new RgbColorModelHex()
            {
                Val = hexString.ToString()
            };

            var outlineNoFill = new Outline();
            outlineNoFill.Append(new NoFill());

            solidFill.Append(hexColour);
            chartShapeProperties.Append(solidFill);
            chartShapeProperties.Append(outlineNoFill);
            marker.Append(chartShapeProperties);
        }

        /// <summary>
        /// Updates the category value chart series.
        /// </summary>
        /// <param name="series">The series.</param>
        /// <param name="newSeriesIndex">New index of the series.</param>
        /// <param name="categoryHeadingRange">The category heading range.</param>
        /// <param name="axisDataRange">The axis data range.</param>
        /// <param name="seriesDataRange">The series data range.</param>
        /// <exception cref="ArgumentNullException">series</exception>
        /// <exception cref="InvalidOperationException">Only valid for series of type BarChartSeries, LineChartSeries, AreaChartSeries & PieChartSeries</exception>
        public static void UpdateCategoryValueChartSeries(this OpenXmlCompositeElement series,
                                                          uint newSeriesIndex,
                                                          CompositeRangeReference categoryHeadingRange,
                                                          CompositeRangeReference axisDataRange,
                                                          CompositeRangeReference seriesDataRange)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            if (!(series is BarChartSeries) && !(series is LineChartSeries) && !(series is AreaChartSeries) && !(series is PieChartSeries))
            {
                throw new InvalidOperationException("Only valid for series of type BarChartSeries, LineChartSeries, AreaChartSeries & PieChartSeries");
            }

            // Updates the supplied series Index and Order
            var index = series.Descendants<Index>().FirstOrDefault();
            if (index != null)
            {
                index.Val = (UInt32Value)newSeriesIndex;
            }

            var order = series.Descendants<Order>().FirstOrDefault();
            if (order != null)
            {
                order.Val = (UInt32Value)newSeriesIndex;
            }

            // Set the formula on the SeriesText ('Legend Entries(Series)' in Excel).
            // This is set to reference the column header in Excel and determines the Category Heading
            var seriesText = series.Descendants<SeriesText>().FirstOrDefault();
            if (seriesText != null)
            {
                seriesText.StringReference = new StringReference();
                seriesText.StringReference.Formula = new Formula();
                seriesText.StringReference.Formula.Text = categoryHeadingRange.Reference;
            }

            // Set the formula on the Category Axis
            var categoryAxisData = series.Descendants<CategoryAxisData>().FirstOrDefault();
            if (categoryAxisData != null)
            {
                categoryAxisData.StringReference = new StringReference();
                categoryAxisData.StringReference.Formula = new Formula();
                categoryAxisData.StringReference.Formula.Text = axisDataRange.Reference;
            }

            // Set the Formula on the data
            var values = series.Descendants<Values>().FirstOrDefault();
            if (values != null)
            {
                values.NumberReference = new NumberReference();
                values.NumberReference.Formula = new Formula();
                values.NumberReference.Formula.Text = seriesDataRange.Reference;
            }
        }

        /// <summary>
        /// Updates the category value chart series.
        /// </summary>
        /// <param name="series">The series.</param>
        /// <param name="isRowSeries">if set to <c>true</c> [is row series].</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="seriesCount">The series count.</param>
        /// <param name="itemPosition">The item position.</param>
        /// <param name="xCount">The x count.</param>
        /// <param name="xOffset">The x offset.</param>
        /// <param name="yCount">The y count.</param>
        /// <param name="yOffset">The y offset.</param>
        /// <exception cref="ArgumentNullException">series</exception>
        /// <exception cref="InvalidOperationException">Only valid for series of type BarChartSeries & LineChartSeries</exception>
        [Obsolete("Use UpdateCategoryValueChartSeries(updated parameter list) instead. Remains for backwards compatibility.!")]
        public static void UpdateCategoryValueChartSeries(this OpenXmlCompositeElement series, 
                                                          bool isRowSeries, 
                                                          string sheetName,
                                                          uint seriesCount,
                                                          uint itemPosition,
                                                          uint xCount,
                                                          uint xOffset,
                                                          uint yCount,
                                                          uint yOffset)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            if (!(series is BarChartSeries) && !(series is LineChartSeries))
            {
                throw new InvalidOperationException("Only valid for series of type BarChartSeries & LineChartSeries");
            }

            // Updates the supplied series Index and Order
            var index = series.Descendants<Index>().FirstOrDefault();
            if (index != null)
            {
                index.Val = (UInt32Value)seriesCount;
            }

            var order = series.Descendants<Order>().FirstOrDefault();
            if (order != null)
            {
                order.Val = (UInt32Value)seriesCount;
            }

            // Set the SeriesText ('Legend Entries(Series)' in Excel).
            // This is set to reference the column header in Excel and determines the Category.
            var seriesText = series.Descendants<SeriesText>().FirstOrDefault();
            if (seriesText != null)
            {
                seriesText.StringReference = new StringReference();
                seriesText.StringReference.Formula = new Formula();

                if (isRowSeries)
                {
                    seriesText.StringReference.Formula.Text = string.Format("'{0}'!$A${1}", sheetName, itemPosition + 1);
                }
                else 
                {
                    // a column is a series
                    seriesText.StringReference.Formula.Text = string.Format("'{0}'!${1}${2}", sheetName, CellExtensions.GetColumnLetter(itemPosition + 1), yOffset);
                }
            }

            // set the formula for the category axis, currently always the 1st column of data 
            var categoryAxisData = series.Descendants<CategoryAxisData>().FirstOrDefault();
            if (categoryAxisData != null)
            {
                categoryAxisData.StringReference = new StringReference();
                categoryAxisData.StringReference.Formula = new Formula();
                if (isRowSeries)
                {
                    // if row is the series, the category are the dynamic columns
                    // so for example B3:F3
                    categoryAxisData.StringReference.Formula.Text = string.Format("'{0}'!${1}${2}:${3}${2}", sheetName, CellExtensions.GetColumnLetter(xOffset + 1), yOffset, CellExtensions.GetColumnLetter(xCount + 1));
                }
                else 
                {
                    // if row isnt series, then its a category, the default behaviour
                    // so assuming 1st column is the category then formula is A4:A10 with 4 being the starting row in the sheet
                    categoryAxisData.StringReference.Formula.Text = string.Format("'{0}'!$A${1}:$A${2}", sheetName, yOffset + 1, yCount + yOffset);
                }
            }

            var values = series.Descendants<Values>().FirstOrDefault();
            if (values != null)
            {
                values.NumberReference = new NumberReference();
                values.NumberReference.Formula = new Formula();
                if (isRowSeries)
                {
                    values.NumberReference.Formula.Text = string.Format("'{0}'!${1}${2}:${3}${2}", sheetName, CellExtensions.GetColumnLetter(xOffset + 1), itemPosition + 1, CellExtensions.GetColumnLetter(xCount + xOffset));
                }
                else 
                {
                    values.NumberReference.Formula.Text = string.Format("'{0}'!${1}${2}:${1}${3}", sheetName, CellExtensions.GetColumnLetter(itemPosition + 1), yOffset + 1, yCount + yOffset);
                }
            }
        }

        /// <summary>
        /// Updates a line/bar series, which is generally cloned from an existing series, so that elements are prepared for insertion into an existing chart.
        /// </summary>
        /// <param name="series">The series that we are updating (must be a <see cref="BarChartSeries" /> or a <see cref="LineChartSeries" /><br />
        /// i.e. a series that consists of a Category to identify the series, and a series value.</param>
        /// <param name="newSeriesIndex">The 0-based index/order to be set for the new series in the charts's Series collection</param>
        /// <param name="categoryHeadingRange">The category heading range.</param>
        /// <param name="axisDataRange">The axis data range.</param>
        /// <param name="seriesDataRange">The series data range.</param>
        /// <exception cref="ArgumentNullException">series</exception>
        /// <exception cref="InvalidOperationException">Only valid for series of type ScatterChartSeries</exception>
        public static void UpdateXYValueChartSeries(this OpenXmlCompositeElement series,
                                                    uint newSeriesIndex,
                                                    CompositeRangeReference categoryHeadingRange,
                                                    CompositeRangeReference axisDataRange,
                                                    CompositeRangeReference seriesDataRange)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            if (!(series is ScatterChartSeries))
            {
                throw new InvalidOperationException("Only valid for series of type ScatterChartSeries");
            }

            var index = series.Descendants<Index>().FirstOrDefault();
            if (index != null)
            {
                index.Val = (UInt32Value)newSeriesIndex;
            }

            var order = series.Descendants<Order>().FirstOrDefault();
            if (order != null)
            {
                order.Val = (UInt32Value)newSeriesIndex;
            }

            // the series title, this is the name of the column header
            var seriesText = series.Descendants<SeriesText>().FirstOrDefault();
            if (seriesText != null)
            {
                seriesText.StringReference = new StringReference();
                seriesText.StringReference.Formula = new Formula();
                seriesText.StringReference.Formula.Text = categoryHeadingRange.Reference;
            }

            // set the formula for the category axis, currently always the 1st column of data 
            var xvalues = series.Descendants<XValues>().FirstOrDefault();
            if (xvalues != null)
            {
                xvalues.StringReference = new StringReference();
                xvalues.StringReference.Formula = new Formula();
                xvalues.StringReference.Formula.Text = axisDataRange.Reference;
            }

            var yvalues = series.Descendants<YValues>().FirstOrDefault();
            if (yvalues != null)
            {
                yvalues.NumberReference = new NumberReference();
                yvalues.NumberReference.Formula = new Formula();
                yvalues.NumberReference.Formula.Text = seriesDataRange.Reference;
            }
        }

        /// <summary>
        /// Updates the xy value chart series.
        /// </summary>
        /// <param name="series">The series.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="column">The column.</param>
        /// <param name="rowOffset">The row offset.</param>
        /// <param name="rowCount">The row count.</param>
        /// <exception cref="ArgumentNullException">series</exception>
        /// <exception cref="InvalidOperationException">Only valid for series of type ScatterChartSeries</exception>
        [Obsolete("Use UpdateXYValueChartSeries(updated parameter list) instead. Remains for backwards compatibility.!")]
        public static void UpdateXYValueChartSeries(this OpenXmlCompositeElement series, string sheetName, uint column, uint rowOffset, uint rowCount)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            if (!(series is ScatterChartSeries))
            {
                throw new InvalidOperationException("Only valid for series of type ScatterChartSeries");
            }

            var index = series.Descendants<Index>().FirstOrDefault();
            if (index != null)
            {
                index.Val = (UInt32Value)column - 1;
            }

            var order = series.Descendants<Order>().FirstOrDefault();
            if (order != null)
            {
                order.Val = (UInt32Value)column - 1;
            }

            // the series title, this is the name of the column header
            var seriesText = series.Descendants<SeriesText>().FirstOrDefault();
            if (seriesText != null)
            {
                seriesText.StringReference = new StringReference();
                seriesText.StringReference.Formula = new Formula();
                seriesText.StringReference.Formula.Text = string.Format("'{0}'!${1}${2}", sheetName, CellExtensions.GetColumnLetter(column + 1), rowOffset);
            }

            // set the formula for the category axis, currently always the 1st column of data 
            var xvalues = series.Descendants<XValues>().FirstOrDefault();
            if (xvalues != null)
            {
                xvalues.StringReference = new StringReference();
                xvalues.StringReference.Formula = new Formula();
                xvalues.StringReference.Formula.Text = string.Format("'{0}'!$A${1}:$A${2}", sheetName, rowOffset + 1, rowCount + rowOffset);
            }

            var yvalues = series.Descendants<YValues>().FirstOrDefault();
            if (yvalues != null)
            {
                yvalues.NumberReference = new NumberReference();
                yvalues.NumberReference.Formula = new Formula();
                yvalues.NumberReference.Formula.Text = string.Format("'{0}'!${1}${2}:${1}${3}", sheetName, CellExtensions.GetColumnLetter(column + 1), rowOffset + 1, rowCount + rowOffset);
            }
        }
    }
}



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using Gam.MM.Framework.Export;

namespace Gam.MM.Framework.Export.Map
{
    /// <summary>
    /// When implemented on a <see cref="DataPart"/> row or <see cref="TableColumn"/>,
    /// this property determines information about the row/column which<br/>
    /// relates to how it is presented in charts.
    /// </summary>
    [Obsolete("This class will be removed, please use IChartMetadata instead, specifying ChartOptions as required")]
    public interface IChartSeries
    {
        /// <summary>
        /// If set to true, then any series generated for charts
        /// off this row/column data, will be suppressed.
        /// </summary>
        [Obsolete("Suppress by using ChartOptions instead.")]
        bool SuppressSeries { get; set; }

        /// <summary>
        /// If set to true, then this row/column will be used as the X1 (or Category) Axis for charts.<br/>
        /// If not set, then the first row/column in the table will be used as the X1 (Category) Axis.
        /// </summary>
        bool IsCategory1Axis { get; set; }

        /// <summary>
        /// Specifies the 0-based index of the series in the template chart on
        /// which this row/column of table data will be based.<br/>
        /// Default value is 0 (ie. 1st series in template chart).
        /// </summary>
        int BaseOnChartSeriesIndex { get; set; }

        /// <summary>
        /// Specifies a brush which can be used to set the colour of the chart series.
        /// </summary>
        [Obsolete("Note that this is not thread-safe and the supplied Brush should not be created on any UI thread.")]
        SolidColorBrush Brush { get; set; }
    }
}

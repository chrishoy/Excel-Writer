// -----------------------------------------------------------------------
// <copyright file="ChartDataRangeInfo.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using OpenXml.Excel;

    /// <summary>
    /// Represents the information that a chart requries in order to map data to a series.
    /// </summary>
    internal class ChartDataRangeInfo
    {
        #region Private Fields

        private CompositeRangeReference categoryHeadingRange;
        private CompositeRangeReference axisDataRange;
        private CompositeRangeReference seriesDataRange;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Data range which is mapped to the series heading (used in Legends)
        /// </summary>
        public CompositeRangeReference SeriesTextRange
        {
            get { return categoryHeadingRange; }
            set { categoryHeadingRange = value; }
        }

        /// <summary>
        /// Data range which is mapped to the axis data
        /// </summary>
        public CompositeRangeReference CategoryAxisDataRange
        {
            get { return axisDataRange; }
            set { axisDataRange = value; }
        }

        /// <summary>
        /// Data range which is mapped to the series data
        /// </summary>
        public CompositeRangeReference SeriesValuesRange
        {
            get { return seriesDataRange; }
            set { seriesDataRange = value; }
        }

        #endregion Public Properies

    }
}

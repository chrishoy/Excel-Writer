using System;
using System.Collections.Generic;

namespace ExcelWriter
{
    using DocumentFormat.OpenXml;
    using OpenXml.Excel.Model;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    internal class SeriesFactory
    {
        #region Private Fields

        private List<TemplateSeriesInfo> seriesInfoList;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Construction
        /// </summary>
        /// <param name="chartModel"></param>
        public SeriesFactory(ChartModel chartModel)
        {
            if (chartModel == null) throw new ArgumentNullException("chartModel");

            // Create a list of information about template series which stores the OpenXML source for the template series,
            // information about it's clone usage, and the number of times it is used.
            this.seriesInfoList = new List<TemplateSeriesInfo>();
            foreach (var seriesElement in chartModel.GetAllSeriesElements())
            {
                var templateSeriesInfo = new TemplateSeriesInfo(seriesElement);
                seriesInfoList.Add(templateSeriesInfo);
            }

        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets the number of series that can be used as template series within the supplied <see cref="ChartModel"/>
        /// </summary>
        public int SourceSeriesCount
        {
            get { return this.seriesInfoList.Count; }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Gets the use-count of a supplied chart series index.
        /// </summary>
        /// <param name="chartSeriesIndex"></param>
        /// <returns></returns>
        public int GetUseCount(int chartSeriesIndex)
        {
            // Make sure supplied index is within bounds
            AssertIndex(chartSeriesIndex);

            // Pluck out the template series and update its use count
            var templateSeriesInfo = this.seriesInfoList[chartSeriesIndex];

            return templateSeriesInfo.UseCount;
        }

        /// Gets the source OpenXML element for the series of a supplied chart series index.
        /// </summary>
        /// <param name="chartSeriesIndex">Index of series</param>
        /// <returns></returns>
        public OpenXmlCompositeElement GetSourceSeriesElement(int chartSeriesIndex)
        {
            // Make sure supplied index is within bounds
            AssertIndex(chartSeriesIndex);

            // Pluck out the template series and update its use count
            var templateSeriesInfo = this.seriesInfoList[chartSeriesIndex];

            return templateSeriesInfo.SourceSeriesElement;
        }

        /// Gets the cloned OpenXML elements for the series of a supplied chart series index.
        /// </summary>
        /// <param name="chartSeriesIndex">Index of series</param>
        /// <returns></returns>
        public IEnumerable<OpenXmlCompositeElement> GetClonedSeriesElements(int chartSeriesIndex)
        {
            // Make sure supplied index is within bounds
            AssertIndex(chartSeriesIndex);

            // Pluck out the template series and update its use count
            var templateSeriesInfo = this.seriesInfoList[chartSeriesIndex];

            return templateSeriesInfo.ClonedSeriesElements;
        }

        /// <summary>
        /// Gets an OpenXML chart series template based on an index.
        /// If it has a use-count of 0 (not used yet), then returns the series itself.<br/>
        /// If it has already been used, then return a clone and updates in internal use-count.
        /// </summary>
        /// <param name="basedOnIndex"></param>
        /// <returns></returns>
        public OpenXmlCompositeElement GetOrCloneSourceSeries(int chartSeriesIndex)
        {
            // Make sure supplied index is within bounds
            AssertIndex(chartSeriesIndex);

            // Pluck out the template series and update its use count
            var templateSeriesInfo = this.seriesInfoList[chartSeriesIndex];
            templateSeriesInfo.UseCount++;

            // Either grab the tempalte series so we can update it, or clone it if already used
            OpenXmlCompositeElement clonedSeries;
            if (templateSeriesInfo.UseCount == 1)
            {
                clonedSeries = templateSeriesInfo.SourceSeriesElement;
            }
            else
            {
                // Clone the tempalte series and track it so we can set its XML Index and Order properties
                clonedSeries = (OpenXmlCompositeElement)templateSeriesInfo.SourceSeriesElement.CloneNode(true);
                var cloneList = templateSeriesInfo.ClonedSeriesElements as List<OpenXmlCompositeElement>;
                cloneList.Add(clonedSeries);
            }

            return clonedSeries;
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Asserts that the supplied index is within the bounds of the currently loaded chart model template series.
        /// </summary>
        /// <param name="basedOnIndex"></param>
        private void AssertIndex(int basedOnIndex)
        {
            // Check valid index
            if (basedOnIndex < 0 || basedOnIndex > this.seriesInfoList.Count - 1)
            {
                throw new InvalidOperationException(string.Format("A series can only be created based on a source series index within the range of '0 to {0}' (the number of series in the template chart - 1). The index requested was {1}", this.seriesInfoList.Count - 1, basedOnIndex));
            }
        }

        #endregion Private Helpers
    }
}

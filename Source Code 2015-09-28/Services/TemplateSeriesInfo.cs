// -----------------------------------------------------------------------
// <copyright file="TemplateSeriesInfo.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml;

    /// <summary>
    /// Represents information about an OpenXml series and its usage for cloning.
    /// </summary>
    internal class TemplateSeriesInfo
    {
        #region Private Fields

        private OpenXmlCompositeElement sourceSeriesElement;
        private IEnumerable<OpenXmlCompositeElement> clonedSeriesElements;
        private int useCount;



        #endregion Private Fields

        #region Construction

        /// <summary>
        ///  Constructor
        /// </summary>
        /// <param name="sourceSeriesElement"></param>
        public TemplateSeriesInfo(OpenXmlCompositeElement sourceSeriesElement)
        {
            this.sourceSeriesElement = sourceSeriesElement;
            this.clonedSeriesElements = new List<OpenXmlCompositeElement>();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets the source OpenXML series element
        /// </summary>
        public OpenXmlCompositeElement SourceSeriesElement
        {
            get { return sourceSeriesElement; }
        }

        /// <summary>
        /// Gets a list of clones.
        /// </summary>
        public IEnumerable<OpenXmlCompositeElement> ClonedSeriesElements
        {
            get { return clonedSeriesElements; }
        }

        /// <summary>
        /// Gets or sets a value which indicates the number of times this template series has been used.
        /// </summary>
        public int UseCount
        {
            get { return useCount; }
            set { useCount = value; }
        }
        
        #endregion Public Properties

        #region Public Methods


        #endregion Public Methods
    }
}

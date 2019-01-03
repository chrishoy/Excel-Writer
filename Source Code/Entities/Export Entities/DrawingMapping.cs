using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    public class DrawingMapping : VisualMapping
    {
        /// <summary>
        /// The name of the drawing to mapping to the PlaceholderId
        /// If not supplied then the 1st chart is taken 
        /// </summary>
        public string SourceDrawingId { get; set; }

        /// <summary>
        /// If supplied create a 'legend' object and position in the supplied placeholder id
        /// The legend will be based on the chart that's being mapped
        /// </summary>
        public string LegendPlaceholderId { get; set; }
    }
}

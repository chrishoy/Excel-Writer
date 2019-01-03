using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter.OpenXml.Excel.Model
{
    /// <summary>
    /// Use for internal puposes only to control the scope of what this code can handel.
    /// If you wish to expose then consider using OpenXML specific types.
    /// </summary>
    public enum SeriesType
    {
        Unrecognised,
        Line,
        //Bar,
        Scatter,
        Pie,
    }
}

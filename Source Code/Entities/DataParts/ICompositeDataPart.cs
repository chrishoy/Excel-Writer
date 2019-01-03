using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Container for multiple data parts
    /// </summary>
    public interface ICompositeDataPart : IDataPart
    {
       IEnumerable<IDataPart> DataParts { get; set; }
    }
}

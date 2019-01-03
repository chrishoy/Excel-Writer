using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Gam.MM.Framework.Export
{
    /// <summary>
    /// 
    /// </summary>
    public class ExportPartCollection
    {
        public ExportPartCollection()
        {
            this.Parts = new List<ExportPart>();
        }

        public ExportPart GetPartById(string partId)
        {
            return (from p in this.Parts
                    where p.PartId.CompareTo(partId) == 0
                    select p).FirstOrDefault();
        }

        public List<ExportPart> Parts { get; private set; }
    }
}

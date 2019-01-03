using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    public sealed class MappingPlaceholderSet
    {
        public MappingPlaceholderSet()
        {
            this.Items = new List<MappingPlaceholderSetItem>();
        }

        public string Id { get; set; }

        /// <summary>
        /// </summary>
        public List<MappingPlaceholderSetItem> Items { get; set; }        
    }
}

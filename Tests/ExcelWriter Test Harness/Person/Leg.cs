using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter.TestHarness.Person
{
    public class Leg
    {
        public Leg()
        {
            this.Foot = new Foot();
        }
        public Foot Foot { get; set; } 
    }
}

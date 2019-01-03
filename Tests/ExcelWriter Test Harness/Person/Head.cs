using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter.TestHarness.Person
{
    public class Head
    {

        public Head()
        {
            this.Eyes = new List<Eye> {new Eye {Colour = "Blue"}, new Eye {Colour = "Blue"}};

            this.Hairs = new List<Hair>();

            int numberOfHairs = 0;
            while (numberOfHairs < 1000)
            {
                numberOfHairs = numberOfHairs + 1;
                this.Hairs.Add(new Hair());
            }
        }
        public List<Eye> Eyes { get; set; }

        public List<Hair> Hairs { get; set; } 
    }
}

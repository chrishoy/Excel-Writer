using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.TestHarness.Person
{
    public enum Gender
    {
        Male,
        Female,
    }

    public class Person
    {
        public Person(Gender gender, string name)
        {
            this.Arms = new List<Arm> {new Arm(), new Arm()};

            this.Legs = new List<Leg> {new Leg(), new Leg()};

            this.Name = name;

            this.Head = new Head();

            this.Mouth = new Mouth();

            this.Gender = gender;
        }

        public Gender Gender { get; set; }

        public Mouth Mouth { get; set; }

        public Head Head { get; set; }

        public List<Arm> Arms { get; set; }

        public List<Leg> Legs { get; set; }

        public string Name { get; set; }
    }
}

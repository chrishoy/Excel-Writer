using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.TestHarness.Person
{
    public class Class
    {
        public Class()
        {
            this.Children = new List<Person>();
            this.Children.Add(new Person(Gender.Female, "Ellie Hoy"));
            this.Children.Add(new Person(Gender.Male, "Callum Bleazard"));
            this.Children.Add(new Person(Gender.Female, "Charlotte Wilkinson"));
            this.Children.Add(new Person(Gender.Male, "Joe Fellows"));
            this.Children.Add(new Person(Gender.Female, "Sofi Selby"));
            this.Children.Add(new Person(Gender.Male, "Jack Dunn"));
            this.Children.Add(new Person(Gender.Female, "Adley Webb"));
            this.Children.Add(new Person(Gender.Male, "Adam Howling"));
            this.Children.Add(new Person(Gender.Female, "Levi Woodland"));
            this.Children.Add(new Person(Gender.Male, "Oliver Farrel"));
            this.Children.Add(new Person(Gender.Female, "Eve McCafery"));
            this.Children.Add(new Person(Gender.Male, "Abdul Khan"));
            this.Children.Add(new Person(Gender.Female, "Leah Harris"));
            this.Children.Add(new Person(Gender.Male, "Freddie Chalk"));
        }

        public List<Person> Children { get; set; } 

    }
}

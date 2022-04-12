using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingExecl.Model
{
    public class Student
    {
        //Properties
        public string Name { get; set; }
        public int Age { get; set; }

        public string Qualification { get; set; }

        public double Height { get; set; }

        public IEnumerable<Student> GetData()
        {
            // Create a list of accounts.
            var TestDetails = new List<Student>
            {
                new Student {Name="Ashiq",Age=32,Qualification="CS",Height=5.6 },
                new Student {Name="Sanil",Age=33,Qualification="MBA",Height=5.7},
                new Student {Name="Manu",Age=25,Qualification="Btech",Height=7.2},
                new Student {Name="Rahul",Age=20,Qualification="BE",Height=5.2},
              new Student {Name="yasir",Age=28,Qualification="ME",Height=6.1}
            };

            return TestDetails;
        }


    }
}


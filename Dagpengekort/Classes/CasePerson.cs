using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dagpengekort.Classes
{
    public class CasePerson
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public double Dagpengeret { get; set; }

        public CasePerson(string name, int age, double dagpengeret) 
        {
            Name = name;    
            Age = age;
            Dagpengeret = dagpengeret;
        }
    }
}

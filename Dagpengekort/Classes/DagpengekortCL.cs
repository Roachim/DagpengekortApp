using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dagpengekort.Classes
{
    public class DagpengekortCL
    {

        public Dictionary<string, string>? Medlem { get; set; }
        public Dictionary<string, string>? Arbejdsgiver { get; set; }
        public Dictionary<string, string>? Akasse { get; set; }
        public Dictionary<string, string>? Header { get; set; }
        public Dagpengespecifikationer? Dagpengespecifikationer { get; set; }
        public Dictionary<string, dynamic>? Footer { get; set; }



        public DagpengekortCL() { }

    }
}

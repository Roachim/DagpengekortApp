using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dagpengekort.Classes
{
    public class Dagpengespecifikationer
    {
        public Dictionary<string, string>? Periode { get; set; }
        public double Timer { get; set; }
        public double Timesats { get; set; }
        public double Dagpengebeløb { get; set; }
        public double ATP_sats { get; set; }
        public double ATP { get; set; }
        public double Bruto_Udbetaling { get; set; }
        public string Trækprocent { get; set; }
        public double Månedsfradrag { get; set; }
        public double Skat { get; set; }
        public double Netto_Udbetaling { get; set; }

        public Dagpengespecifikationer() { }
    }
}

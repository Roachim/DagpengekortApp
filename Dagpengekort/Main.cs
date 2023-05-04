using Dagpengekort.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using Excel = Microsoft.Office.Interop.Excel;   //A COM reference to handle the excel file
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Net.Http.Json;
using System.Xml;
using System.Text.Encodings.Web;
using Microsoft.Office.Interop.Excel;

namespace Dagpengekort
{
    public static class Main
    {
        //private static string TestFilePath = "C:\\Users\\KOM\\Desktop\\Opgaver\\Dagpenge kort\\Dagpengekort til opgave.xlsx";
        //private static string TestFilePath2 = "..\\..\\..\\..\\";
        public static void Run()
        {
            CreateJSONFolder();

            string filePath = DagpengekortAbsolutePath();
            //Open the excel app
            Excel.Application xlApp = new Excel.Application();
            //Get the excel file
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);

            //how to run through different pages of same excel sheet

            foreach (Excel.Worksheet ws in xlWorkbook.Sheets) //run through every sheet
            {
                Excel.Range wsRange = ws.UsedRange; //get the used range for every sheet run through

                int rows = wsRange.Rows.Count;      // Setting counters outside the loop speeds it up
                int cols = wsRange.Columns.Count;
                ReadExSheet(wsRange, rows, cols);

                Marshal.ReleaseComObject(ws.UsedRange);
                Marshal.ReleaseComObject(ws);

            }



            //lastly Cleanup - This is important: To prevent lingering processes from holding the file access writes to the workbook
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        /// <summary>
        /// finds the absolute path to the excel file used by the program to generate JSON files with.
        /// </summary>
        /// <returns>string path to excel file</returns>
        private static string DagpengekortAbsolutePath()
        {
            var appFolder = System.IO.Directory.GetDirectories("..\\..\\..\\..\\..");
            string val = "";
            foreach (var str in appFolder)
            {
                if (str.EndsWith("DagpengekortApp"))
                    val += str;
            }
            val = val + "\\Dagpengekort til opgave.xlsx";
            string path1 = Path.GetFullPath(val);
            return path1;
        }

        /// <summary>
        /// Creates a folder for JSON files
        /// </summary>
        private static void CreateJSONFolder()
        {
            string val = JSONFolderAbsolutePath();

            if (System.IO.Directory.Exists(val)) //Delete and create a download folder for files to be downloaded
            {
                System.IO.Directory.Delete(val, true);   //True to delete everything within the folder as well
            }
            else
            {
                System.IO.Directory.CreateDirectory(val);
            }
        }

        /// <summary>
        /// Finds the absolute path to the folder created by this program
        /// </summary>
        /// <returns>string</returns>
        private static string JSONFolderAbsolutePath()
        {
            var appFolder = System.IO.Directory.GetDirectories("..\\..\\..\\..\\..");
            string val = "";
            foreach (var str in appFolder)
            {
                if (str.EndsWith("DagpengekortApp"))
                    val += str;
            }
            val = val + "\\JSONFiles";
            return val;
        }

        /// <summary>
        /// Read a single sheet in an excel file.
        /// Calls Create Json
        /// </summary>
        /// <param name="wsRange">Given range for the sheet</param>
        /// <param name="rows">Amount of used rows in the sheet</param>
        /// <param name="cols">Amount of used columns in the sheet</param>
        private static void ReadExSheet(Excel.Range wsRange, int rows, int cols)
        {
            List<string> importantWords = new List<string>
            {
                "Teknisk belægning",
                "Ferie",
                "Arbejdstimer",
                "G-dage",
                "Sygdom"
            };

            double monthTotal = 0;

            DateTime? modtageDato = null;
            DateOnly dateOnly;
            

            int LastDay = 0;
            double lastPayday = 0;

            bool arbejde = false;
            
            // running through every item in the sheet

            for (int i = 1; i <= rows; i++)     //i = 1. excel does not start at 0.
            {
                
                for (int j = 1; j <= cols; j++)
                {
                    
                    if (wsRange.Cells[i, j].Value2 is null) { continue; }   //make sure there is something here

                    if (wsRange.Cells[i, 1].Value2.ToString() == "Måned" && i != 1) { monthTotal = 0; } //if we see month again, then we have reached a new dagpengekort, and should reset

                    if (wsRange.Cells[i, j].Value2.ToString() != "Fradrag pr. dag" && wsRange.Cells[i, 1].Value2.ToString() == "Fradrag pr. dag")
                    {
                        monthTotal += wsRange.Cells[i, j].Value2;
                    }
                    if(modtageDato == null && wsRange.Cells[i, j].Value2.ToString() != "Modtagedato" && wsRange.Cells[i, 1].Value2.ToString() == "Modtagedato")
                    {
                        dateOnly = DateOnly.FromDateTime(DateTime.FromOADate(wsRange.Cells[i, j].Value2)); 
                    }
                    if(wsRange.Cells[i, 1].Value2 == "Dato" && wsRange.Cells[i, j+1].Value2 is null) // it is assumed that every column on row 3 is not empty, unless there are no more dates left
                    {
                        LastDay = (int)wsRange.Cells[i, j].Value2;

                        bool yet = true;
                        int pol = 0;
                        while (yet)
                        {
                            //from last day of month, go one back until we find the start of a week (which will be a monday).
                            //From there, take up to, but no more than, 4 steps. Starting from monday and reaching friday
                            //This is a rather simple implementation that does not take vacation into account

                            if (wsRange.Cells[2, j - pol].Value2 is not null)
                            {
                                for (int q = 4; q > 0; q--)
                                {
                                    if (wsRange.Cells[i, j - pol + q].Value2 is not null)
                                    {
                                        lastPayday = wsRange.Cells[i, j - pol + q].Value2;
                                        break;
                                    }
                                }
                                yet = false;
                            }
                            pol++;
                        }
                    }

                    //find day to transfer money

                    if (j != 1 && wsRange.Cells[i, 1].Value2.ToString() == "Arbejdstimer") { arbejde = true; }
                }
            }
            DateOnly SlutDato = new DateOnly(dateOnly.Year, dateOnly.Month, LastDay);
            DateOnly StartDato = new DateOnly(dateOnly.Year, dateOnly.Month, (int)wsRange.Cells[3, 2].Value2);
            DateOnly DispositionsDato = new DateOnly(dateOnly.Year, dateOnly.Month, (int)lastPayday);


            //Console.WriteLine(monthTotal + " is the total withdrawn hours");  //reduced hours from the month
            //Console.WriteLine(dateOnly + " is the date it was sent");    //date received
            //Console.WriteLine(StartDato+ " is the first date of the month");    //first date
            //Console.WriteLine(SlutDato+ " is the last date of the month");    //last date

            CreateJSON(monthTotal, dateOnly, StartDato, SlutDato, DispositionsDato, arbejde);
        }

        /// <summary>
        /// Creates a JSON file in a given folder
        /// </summary>
        /// <param name="withdrawnHours">Hours registered on a dagpengekort</param>
        /// <param name="Udskrivningsdato">Date the dagpengekort was sent by the person</param>
        /// <param name="startPeriode">The first day of the month</param>
        /// <param name="slutPeriode">The last day of the month</param>
        /// <param name="dispositionsDato">Day for the money to be distributed to the person</param>
        /// <param name="arbejde">Did they work this month?</param>
        private static void CreateJSON(double withdrawnHours, DateOnly Udskrivningsdato, DateOnly startPeriode, DateOnly slutPeriode, DateOnly dispositionsDato, bool arbejde)
        {

            CasePerson person = new CasePerson("Palle Jensen", 40, 19728);  //creating/giving a case person could be its own method

            double timesats = person.Dagpengeret / 160.33;
            timesats = Math.Round(timesats, 2);
            double timer = 160.33 - withdrawnHours;
            timer = Math.Round(timer, 2);

            double DagpengeBeløb = timer * timesats;
            DagpengeBeløb = Math.Round(DagpengeBeløb, 2);

            double ATPSats = 1.34;

            double ATP = timer * ATPSats;
            ATP = Math.Round(ATP, 2);

            double BrutoUdbetaling = DagpengeBeløb - ATP;
            double Trækprocent = 0.37;
            double Månedsfradrag = 8607;

            double Skat = (BrutoUdbetaling - Månedsfradrag) * Trækprocent;
            Skat = Math.Round(Skat, 2);

            double NettoUdbetaling = BrutoUdbetaling - Skat;

            if(timer < 14.8)
            {
                DagpengeBeløb = 0;
                BrutoUdbetaling = 0;
                NettoUdbetaling = 0;
            }


            string ArGiNavn = "";
            string ArGiAdresse = "";
            string ArGiPostnummer = "";
            string ArGiBy = "";
            string ArGiCVRNummer = "";

            if(arbejde)
            {
                 ArGiNavn = "Specialister";
                 ArGiAdresse = "Lautruphøj 1A";
                 ArGiPostnummer = "2750";
                 ArGiBy = "Ballerup";
                 ArGiCVRNummer = "27351034";
            }


            var dagpengekort = new DagpengekortCL
            {
                Medlem = new Dictionary<string, string>
                {
                    ["Navn"] = "Palle Jensen",
                    ["Adresse"] = "Bybev 25",
                    ["Postnummer"] = "6000",
                    ["By"] = "Kolding"
                },Arbejdsgiver = new Dictionary<string, string>     //udfyldes for måneden, hvis personen har arbejdet  //some sort of bool here i guess
                {
                    ["Navn"] = ArGiNavn,
                    ["Adresse"] = ArGiAdresse,
                    ["Postnummer"] = ArGiPostnummer,
                    ["By"] = ArGiBy,
                    ["CVR-nummer"] = ArGiCVRNummer
                },Akasse = new Dictionary<string, string>
                {
                    ["Navn"] = "Maistrenes AKasse",
                    ["Adresse"] = "Peter Bangs vej 30",
                    ["Postnummer"] = "2000",
                    ["By"] = "Frederiksberg"
                },Header = new Dictionary<string, string>
                {
                    ["Ydelse"] = "Dagpenge",
                    ["Medlemsnummer"] = "12345678",
                    ["Personnummer"] = "2501852117",
                    ["Udskrivningsdato"] = Udskrivningsdato.ToString()       //udfyles auto  //dagen dette udskrives? nej, det er nok 
                },Dagpengespecifikationer = new Dagpengespecifikationer
                {
                    Periode = new Dictionary<string, string>
                    {
                        ["Startdato"] = startPeriode.ToString(), //udfyles auto  //måned start?
                        ["Slutdato"] = slutPeriode.ToString()   //udfyles auto  //måned slut?
                    },
                    Timer = timer,      //udfyles auto  //Udregn fra dagpengekortet
                    Timesats = timesats,   //udfyles auto  //Beregnes
                    Dagpengebeløb = DagpengeBeløb,  //udfyles auto  //Timer x timesats
                    ATP_sats = ATPSats,
                    ATP = ATP,    //udfyles auto  //Timer x ATP-sats
                    Bruto_Udbetaling = Math.Round(BrutoUdbetaling, 2) ,   //udfyles auto  //dagpengebeløb - ATP
                    Trækprocent = "37%",
                    Månedsfradrag = 8607,
                    Skat = Skat,              //udfyles auto   //(Brutto udbetaling - månedsfradrag) * 37%  
                    Netto_Udbetaling = Math.Round(NettoUdbetaling, 2)       //udfyles auto    //Brutto udbetaling - skat  
                },
                Footer = new Dictionary<string, dynamic>
                {
                    ["Indsat på konto"] = "Nem konto",
                    ["Dispositionsdato"] = dispositionsDato.ToString(),      //udfyles auto  //dagen pengene overføres?
                    ["Til udbetaling"] = Math.Round(NettoUdbetaling, 2) ,      //udfyles auto  //samme som Netto udbetaling
                }
            };


            var options = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping ,WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(dagpengekort, options);

            Console.WriteLine(jsonString);

            //string fileplace = "..\\..\\..\\..\\";
            string fileplace = JSONFolderAbsolutePath();

            System.IO.File.WriteAllText(fileplace +"\\"+ Udskrivningsdato.ToString() + "utf8.json", jsonString, Encoding.UTF8);
        }
    }
}

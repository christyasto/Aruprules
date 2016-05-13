using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            //           Console.WriteLine(Rulebook.CheckRequiredFields());
            //           Console.WriteLine(Rulebook.CheckConsistencyWithRevit());
            //           Console.WriteLine(Rulebook.CheckRebarArea());
            //           Console.WriteLine(Sources.beam);
            //           Console.WriteLine(Sources.GetRebarAreaByNumberAndDiameter("8H25"));
            string input = "";
            while (input != "exit")
            {
                input = Console.ReadLine();
                if (input == "exit") { break; }
                else {
                    //Match match = Regex.Match(input, "[0-9]{1,}[H][0-9]{1,}[-]{0,1}[0-9]{0,}[(]{0,1}[B|H]{0,1}[)]{0,1}");
                    Console.WriteLine(Sources.RebarDescriptions(input));
                }
            }
            //            Console.WriteLine(Sources.DecodeRebarMark("99H12-200(B)+5H7+6H7 - 200"));
            //           Console.WriteLine(Sources.beam);
        }
    }
}

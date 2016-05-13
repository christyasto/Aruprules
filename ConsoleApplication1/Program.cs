using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

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
            Console.WriteLine(Sources.ExtractRebarLayers("8H12 - 200(B)+5H10+5H16-200+8H16+5H25-200(B)+8H32-300+10H40+2H32+4H25-400(B)"));
            Console.WriteLine(Sources.RebarDescriptions("10H40+2H32+4H25-400(B)") + "\n" + "\n");
            Console.WriteLine(Sources.BeamMarkDescriptions("10B020aPT"));
            //            Console.WriteLine(Sources.DecodeRebarMark("99H12-200(B)+5H7+6H7 - 200"));
            //           Console.WriteLine(Sources.beam);
        }
    }
}

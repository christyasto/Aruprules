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
            Console.WriteLine(Rulebook.Rule1("a", "b"));
            Console.WriteLine(Rulebook.Rule2("b", "b"));
            Console.WriteLine(Rulebook.Rule3(50, 1, 1));
            Console.WriteLine(Sources.GetBeamDepth("2000/1600 x 400"));
        }
    }
}

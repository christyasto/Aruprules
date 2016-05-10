using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class Rulebook
    {
        public static string Rule3(double f_ck, int colNumLo, int colNumHi)
        {
            List<string> fortesting = new List<string>();
            if (f_ck > 50) { return "The variable f_ck should not exceed 50 MPa. Please check your setting!"; }
            else {
                int rowNum = 4;     //------------------------------------------------------Change this into rowNum required
                int LastRow = 4;    //------------------------------------------------------Change this into rowNum required
                for (int i = rowNum; i <= LastRow; i++)
                {
                    for (int colNum = colNumLo; colNum <= colNumHi; colNum++)
                    {
                        string beamType = "L_End (R)";    //-----------------------------------------Change this into beamType in excel
                        if (Sources.GetDefinition(beamType, colNum) == "1")
                        {
                            string userDefinedSectionSize = "2000/1500 x 300"; //------------------Change this into userDefine in excel
                            double beamDepth = Sources.GetBeamDepth(userDefinedSectionSize);
                            double beamWidth = Sources.GetBeamWidth(userDefinedSectionSize);
                            double A_c = beamDepth * beamWidth;
                            double A_min = Sources.GetMinRebarArea(A_c);

                            string cellToCheck = "1"; //----------------------------------------need fix to check if "1" in excel
                            if (cellToCheck.All(char.IsDigit))
                            {
                                if (Int32.Parse(cellToCheck) < A_min)
                                {
                                    fortesting.Add("Rule 3: Should exceed min. rebar area."); //--change into append comment in excel
                                }
                            }
                        }
                    }
                }
                return fortesting[0];
            }
        }
        public static bool Rule1(string a, string b)
        {
            return a == b;
        }
        public static bool Rule2(string a, string b)
        {
            return a == b;
        }
    }
}

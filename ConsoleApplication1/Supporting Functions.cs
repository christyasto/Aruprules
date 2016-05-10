using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class Sources
    {
        public static string GetDefinition(string beamType, int colNum)
        {
            string detailBeamTypes = "L_End (R)"; //----------------------------------------------- Change into range in excel
            string rebarLocations = "1"; //----------------------------------------------- Change into range in excel
            int i = 1;
            while (i <= 13)
            {
                if (detailBeamTypes == beamType)
                {
                    return rebarLocations; //---------------------------------------------Change into range in excel
                }
                i++;
            }
            return null;
        }

        public static double GetMinRebarArea(double A_c)
        {
            double f_ctm = 0.0, f_yk, f_ck, f_cm, area1, area2;
            f_ck = 50.0; //--------------------------------------------------------Change into values in excel
            f_yk = 500.0; //--------------------------------------------------------Change into values in excel

            if (f_ck <= 50)
            {
                f_ctm = 0.3 * Math.Pow(f_ck, (2.0 / 3.0));
            }
            else
            {
                // The formula 2.12 * ln (1+(f_cm/10)) is not implemented for now
            }
            area1 = .0013 * A_c;
            area2 = .26 * f_ctm * A_c / f_yk;

            if (area1 > area2) { return area1; }
            else return area2;
        }

        public static double GetBeamDepth(string section)
        {
            int indexOfLastSlash = section.LastIndexOf("/");
            int indexOfX = section.IndexOf("x");

            if (indexOfLastSlash > 0)
            {
                string result = section.Substring(indexOfLastSlash + 1, indexOfX - indexOfLastSlash - 2);
                return Convert.ToDouble(result);
            }
            else
            {
                string result = section.Substring(0, indexOfX - 1);
                return Convert.ToDouble(result);
            }
        }

        public static double GetBeamWidth(string section)
        {
            int indexOfX = section.IndexOf("x");
            if (indexOfX > 0)
            {
                string result = section.Substring(indexOfX + 1, section.Length - indexOfX - 1);
                return Convert.ToDouble(result);
            }
            else return 0.0;
        }
    }
}

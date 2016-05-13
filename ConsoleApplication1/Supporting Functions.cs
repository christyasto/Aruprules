using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ConsoleApplication1
{
    class Sources
    {
        public static Dictionary<string, List<string>> BeamData = new Dictionary<string, List<string>>();
        public static Dictionary<string, List<string>> RebarData = new Dictionary<string, List<string>>();

        public static string GetDefinition(string beamType, int colNum)
        {
            string detailBeamTypes = "L_End (R)";   //----------------------------------------------- Change into range in excel
            string rebarLocations = "1";            //----------------------------------------------- Change into range in excel
            int i = 1;
            while (i <= 13)
            {
                if (detailBeamTypes == beamType)
                {
                    return rebarLocations;          //---------------------------------------------Change into range in excel
                }
                i++;
            }
            return null;
        }

        public static double GetMinRebarArea(double A_c)
        {
            double f_ctm = 0.0, f_yk, f_ck, area1, area2;
            // double f_cm;
            f_ck = 50.0;        //--------------------------------------------------------Change into values from excel
            f_yk = 500.0;       //--------------------------------------------------------Change into values from excel

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

        public static string LastRow()      //-------------------------------------------- Change into int type
        {
            return null;                    //-------------------------------------------- change into a code that extracts the used range
        }

        public static double GetRebarAreaByNumberAndDiameter(string numberAndDiameter)
        {
            Match match1 = Regex.Match(numberAndDiameter, "(^[0-9]{1,2})([H])([0-9]{1,2})([+| ]{1,3})([0-9]{1,2})([H])([0-9]{1,2})([/D]{0,3})([/D/w]{0,6})");
            Match match2 = Regex.Match(numberAndDiameter, "(^[0-9]{1,2})([H])([0-9]{1,2})([/D]{0,3})([/D/w]{0,6})");
            if (match1.Success)
            {
                int numRebar1 = Convert.ToInt32(match1.Groups[1].Value);
                string diameter1 = match1.Groups[3].Value;
                int numRebar2 = Convert.ToInt32(match1.Groups[5].Value);
                string diameter2 = match1.Groups[7].Value;
                double unitArea1 = GetRebarAreaByDiameter(diameter1);
                double unitArea2 = GetRebarAreaByDiameter(diameter2);

                return numRebar1 * unitArea1 + numRebar2 * unitArea2;
                //return match1.Groups[1].Value + " " + diameter1 + " " + numRebar2 + " " + diameter2; //------------for checking
            }
            else if (match2.Success)
            {
                int numRebar = Convert.ToInt32(match2.Groups[1].Value);
                string diameter = match2.Groups[3].Value;
                double unitArea = GetRebarAreaByDiameter(diameter);

                return numRebar * unitArea;
                //return match2.Groups[1].Value + " " + diameter;   //------------for checking
            }
            else return 0;
        }

        public static double GetRebarAreaByDiameter(string diameter)
        {
            return Math.Round(Math.PI * Math.Pow((Convert.ToInt32(diameter) / 2), 2), 2);
        }

        public static string ExtractRebarLayers(string mark)
        {
            string fixedmark = fixstr(mark);
            MatchCollection match = Regex.Matches(fixedmark, "([0-9]{0,2}[H][0-9]{0,2}[+]{0,1}){0,9}([-][0-9]{1,5})([(][B|H][)]){0,1}");
            bool check = true;

            string layers = "";
            for (int checker = 0; checker < match.Count; checker++)
            { if (match[checker].Success != true) { check = false; } }
            if (check != false)
            {
                for (int layer = 0; layer < match.Count; layer++)
                {
                    layers += $"Layer {layer + 1}".PadRight(8) + ": " + match[layer].Groups[0].Value + "\n";
                }
                return layers;
            }
            else return "input unrecognized.";

        }

        public static string fixstr(string mark)
        {
            int index = 0;
            char[] result = new char[mark.Length];
            for (int i = 0; i < mark.Length; i++)
            {
                if (mark[i] != ' ')
                {
                    result[index++] = char.ToUpper(mark[i]);
                }
            }
            return new string(result, 0, index);
        }

        public static string RebarDescriptions(string mark)
        {
            string fixedmark = fixstr(mark);
            MatchCollection match = Regex.Matches(fixedmark, "([0-9]{0,})([H])([0-9]{0,})([-]{0,1})([0-9]{0,})([(][B|H][)]){0,1}");

            double multiplier = 1;
            string bundle = "";
            if (match[match.Count - 1].Groups[6].Value.IndexOf("(B)") != -1) { multiplier = 2; bundle = " (B)"; }
            else if (match[match.Count - 1].Groups[6].Value.IndexOf("(H)") != -1) { bundle = " (H)"; }

            Dictionary<string, List<string>> data = new Dictionary<string, List<string>>();

            string layers = "";
            layers += "---------- " + $"Descriptions of {mark}" + " ----------" + "\n";


            layers += "\n" + "Number of bar types ".PadRight(20) + "= " + match.Count + "\n"
                + "Spacing ".PadRight(20) + "= " + match[match.Count - 1].Groups[5].Value + "\n" + "\n";
            for (int layer = 0; layer < match.Count; layer++)
            {
                List<string> subdata = new List<string>();

                Double NoBar = Convert.ToDouble(match[layer].Groups[1].Value) * multiplier;
                Double diam = Convert.ToDouble(match[layer].Groups[3].Value);
                Double area = GetRebarAreaByNumberAndDiameter(match[layer].Groups[0].Value) * multiplier;

                //subdata.Add();

                layers += $"Bar Type {layer + 1}".PadRight(4) + ": " + (match[layer].Groups[1].Value + match[layer].Groups[2].Value
                    + match[layer].Groups[3].Value + bundle).PadRight(10) + "_______________" + "\n"
                    + "   |                                 |" + "\n"
                    + "   | Number of bars = " + Convert.ToString(NoBar).PadRight(15) + "|" + "\n"
                    + "   | Diameter       = " + (diam + " mm").PadRight(15) + "|" + "\n"
                    + "   | Total Area     = " + (area + " mm2").PadRight(15) + "|" + "\n"
                    + "   |_________________________________|" + "\n" + "\n";

            }

            layers += "-----------------------------end-----------------------------";
            return layers;
        }

        public static string BeamMarkDescriptions(string mark)
        {
            string fixedmark = fixstr(mark);
            Match match = Regex.Match(fixedmark, "^([0-9]{0,})([B])([0-9]{0,})([A-Z]{0,})([PR][TC])$");
            if (match.Success)
            {
                string level = match.Groups[1].Value;
                string elemen = match.Groups[2].Value;
                string beamNo = match.Groups[3].Value + char.ToLower(Convert.ToChar(match.Groups[4].Value));
                string Dtype = match.Groups[5].Value;

                string layers = $" -----------  Beam {mark}  ----------- " + "\n"
                    + "|    Level".PadRight(20) + "= " + level + "\n"
                    + "|    Element Type".PadRight(20) + "= " + elemen + "\n"
                    + "|    Beam No.".PadRight(20) + "= " + beamNo + "\n"
                    + "|    Dtype".PadRight(20) + "= " + Dtype + "\n"
                    + " ---------------------------------------- ";
                return layers;
            }
            else return "Beam mark cannot be recognized.";
        }

    }

}

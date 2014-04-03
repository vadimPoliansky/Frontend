using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text.RegularExpressions;
using IndInv.Models;

namespace IndInv.Helpers
{
    public class Colour
    {
        public static String getNum(String str)
        {
            if (str == null) { return null; }
            int n;
            int minPos, maxPos;
            string numStr;

            minPos = str.Length - 1;
            maxPos = 0;
            numStr = null;

            if (int.TryParse(str,out n))
            {
                return str;
            }


            for (int i = 1; i <= str.Length; i++ )
            {
                if (int.TryParse(str[i-1].ToString(), out n) || str[i-1] == '-')
                {
                    if (i-1 < minPos) { minPos = i - 1; }
                    if (i-1 > maxPos) { maxPos = i - 1; }
                }
            }

            if (maxPos != 0){
                numStr = str.Substring(minPos, maxPos - minPos + 1).Replace(",", "").Replace("$", "");
            }
            //string num = numStr;
             return numStr;
        }

        public static String getColour(String inStr, String inTarget, String inMonth, String inCustom, Boolean isYTD, Indicators inIndicator)
        {
            string str;
            float outScore;
            string Target = inTarget;

//                if (test == "MH&A ED revisits to own facility within 30 days (%)")
//                    System.Diagnostics.Debugger.Break();

            str = Colour.getNum(inStr);
            if (str == null) { return "cssWhite"; }
            if (float.TryParse(str, out outScore)) { }
            else if (str != null && str.Length != 0 && str.IndexOf(":") != -1) { outScore = float.Parse(str.Replace(":", "")); }
            float score = outScore;

            str = Colour.getNum(Target);
            if (str == null) { return "cssWhite"; }
            if (float.TryParse(str, out outScore)) { }
            else if (str != null && str.Length != 0 && str.IndexOf(":") != -1) { outScore = float.Parse(str.Replace(":", "")); }
            float targetScore = outScore;

            if (inIndicator.Colour_ID == 3){ return inCustom;}

            if (inIndicator.Colour_ID == 2)
            {
                if (isYTD == false)
                {
                    switch (inMonth)
                    {
                        case "Apr": score *= 12 / 1; break;
                        case "May": score *= 12 / 2; break;
                        case "Jun": score *= 12 / 3; break;
                        case "Jul": score *= 12 / 1; break;
                        case "Aug": score *= 12 / 2; break;
                        case "Sep": score *= 12 / 3; break;
                        case "Oct": score *= 12 / 1; break;
                        case "Nov": score *= 12 / 2; break;
                        case "Dec": score *= 12 / 3; break;
                        case "Jan": score *= 12 / 1; break;
                        case "Feb": score *= 12 / 2; break;
                        case "Mar": score *= 12 / 3; break;
                        default: break;
                    }
                }
                else
                {
                    int numMonths = 0;
                    switch (inIndicator.FY_13_14_Q1_Sup)
                    {
                        case "Apr": numMonths+=1; break;
                        case "May": numMonths += 2; break;
                        case "Jun": numMonths += 3; break;
                        default: numMonths += 3; break;
                    }
                    switch (inIndicator.FY_13_14_Q2_Sup)
                    {
                        case "Jul": numMonths += 1; break;
                        case "Aug": numMonths += 2; break;
                        case "Sep": numMonths += 3; break;
                        default: numMonths += 3; break;
                    }
                    switch (inIndicator.FY_13_14_Q3_Sup)
                    {
                        case "Oct": numMonths += 1; break;
                        case "Nov": numMonths += 2; break;
                        case "Dec": numMonths += 3; break;
                        default: numMonths += 3; break;
                    }
                    switch (inIndicator.FY_13_14_Q4_Sup)
                    {
                        case "Jan": numMonths += 1; break;
                        case "Feb": numMonths += 2; break;
                        case "Mar": numMonths += 3; break;
                        default: numMonths += 3; break;
                    }
                    score *= 12 / numMonths;
                }
            }
 
            if (Target != null)
            {
                if (Target.IndexOf("≤") != -1)
                {
                    if (score <= targetScore)
                    {
                        return "cssGreen";
                    }
                    else
                    {
                        if (score <= targetScore * 1.1)
                        {
                            return "cssYellow";
                        }
                        else
                        {
                            return "cssRed";
                        }
                    }
                }
                else if (Target.IndexOf("<") != -1)
                {
                    if (score < targetScore)
                    {
                        return "cssGreen";
                    }
                    else
                    {
                        if (score < targetScore * 1.1)
                        {
                            return "cssYellow";
                        }
                        else
                        {
                            return "cssRed";
                        }
                    }
                }
                else if (Target.IndexOf(">") != -1)
                {
                    if (score > targetScore)
                    {
                        return "cssGreen";
                    }
                    else
                    {
                        if (score > targetScore * 0.9)
                        {
                            return "cssYellow";
                        }
                        else
                        {
                            return "cssRed";
                        }
                    }
                }
                else
                {
                    if (score >= targetScore)
                    {
                        return "cssGreen";
                    }
                    else
                    {
                        if (score >= targetScore * 0.9)
                        {
                            return "cssYellow";
                        }
                        else
                        {
                            return "cssRed";
                        }
                    }
                }
            }
            else
            {
                return "cssWhite";
            }
        }
    }
}
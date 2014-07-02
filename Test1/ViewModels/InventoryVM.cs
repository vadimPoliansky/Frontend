using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class InventoryViewModel
    {
        public string Indicator_ID { get; set; }
        public Int16 Area_ID { get; set; }
        public string CoE {get; set; }
        public string Indicator { get; set; }
        public string FY_3 { get; set; }
        public string FY_3_Sup { get; set; }
        public string FY_2 { get; set; }
        public string FY_2_Sup { get; set; }
        public string FY_1 { get; set; }
        public string FY_1_Sup { get; set; }
        public string FY_Q1 { get; set; }
        public string FY_Q1_Sup { get; set; }
        public string FY_Q2 { get; set; }
        public string FY_Q2_Sup { get; set; }
        public string FY_Q3 { get; set; }
        public string FY_Q3_Sup { get; set; }
        public string FY_Q4 { get; set; }
        public string FY_Q4_Sup { get; set; }
        public string FY_YTD { get; set; }
        public string FY_YTD_Sup { get; set; }
        public string Target { get; set; }
        public string Target_Sup { get; set; }
        public string Comparator { get; set; }
        public string Comparator_Sup { get; set; }
        public string Performance_Threshold { get; set; }
        public string Performance_Threshold_Sup { get; set; }

        public Int16 Colour_ID { get; set; }
        public string Custom_YTD { get; set; }
        public string Custom_Q1 { get; set; }
        public string Custom_Q2 { get; set; }
        public string Custom_Q3 { get; set; }
        public string Custom_Q4 { get; set; }

        public string Definition_Calculation { get; set; }
        public string Target_Rationale { get; set; }
        public string Comparator_Source { get; set; }

        public string Data_Source_MSH { get; set; }
        public string Data_Source_Benchmark { get; set; }
        public string OPEO_Lead { get; set; }

        public string Q1_Color { get; set; }
        public string Q2_Color { get; set; }
        public string Q3_Color { get; set; }
        public string Q4_Color { get; set; }
        public string YTD_Color { get; set; }

        public int Fiscal_Year { get; set; }

        public List<Analysts> allAnalysts { get; set; }
    }
}
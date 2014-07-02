using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class AnalystViewModel
    {
        public Int16 Analyst_ID { get; set; }
        public string Last_Name { get; set; }
        public string First_Name { get; set; }
        public string Position { get; set; }
        public Int16 Order { get; set; }
    }
}
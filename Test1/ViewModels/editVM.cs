using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using IndInv.Models;

namespace IndInv.Models.ViewModels
{
    public class editViewModel
    {
        public Indicators Indicator { get; set; }
        public List<CoEs> allCoEs { get; set; }
    }
}
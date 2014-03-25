using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class Indicator_CoE_MapsViewModel
    {
        public List<Indicators> allIndicators { get; set; }
        public List<CoEs> allCoEs { get; set; }
        public List<Indicator_CoE_Maps> allMaps { get; set; }
    }
}
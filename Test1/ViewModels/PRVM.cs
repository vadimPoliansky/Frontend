using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class PRViewModel
    {
        public List<CoEs> allCoEs { get; set; }
        public List<Indicator_CoE_Maps> allMaps { get; set; }
        public List<Indicator_Footnote_Maps> allFootnoteMaps { get; set; }
    }
}
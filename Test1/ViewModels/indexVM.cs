using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class indexViewModel
    {
        public List<Indicators> allIndicators { get; set; }

        public List<Indicator_Footnote_Maps> allFootnotes { get; set; }
        public List<CoEs> allCoEs { get; set; }
        public List<Areas> allAreas { get; set; }
    }
}
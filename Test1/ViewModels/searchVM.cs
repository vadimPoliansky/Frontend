using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class CoE_IDsViewModel
    {
        public Int16 CoE_ID { get; set; }
    }

    public class Area_IDsViewModel
    {
        public Int16 Area_ID { get; set; }
    }

    public class Indicator_TypesViewModel
    {
        public string Indicator_Type { get; set; }
    }

    public class searchViewModel
    {
        public string searchString { get; set; }
        public List<CoE_IDsViewModel> selectedCoEs { get; set; }
        public List<Area_IDsViewModel> selectedAreas { get; set; }
        public List<Indicator_TypesViewModel> selectedTypes { get; set; }
    }
}
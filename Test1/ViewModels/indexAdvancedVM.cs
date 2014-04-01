using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class indexAdvancedViewModel
    {
        public string searchString { get; set; }
        public List<CoEs> selectedCoEs { get; set; }
    }
}
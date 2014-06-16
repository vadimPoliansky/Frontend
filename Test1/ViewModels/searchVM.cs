using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IndInv.Models.ViewModels
{
    public class selectedCoEs
    {
        public Int16 CoE_ID { get; set; }
    }

    public class selectedAreas
    {
        public Int16 Area_ID { get; set; }
    }

    public class selectedTypes
    {
        public string Indicator_Type { get; set; }
    }

    public class selectedFootnotes
    {
        public Int16 Footnote_ID { get; set; }
    }

    public class searchViewModel
    {
        public string searchString { get; set; }
        public List<selectedCoEs> selectedCoEs { get; set; }
        public List<selectedAreas> selectedAreas { get; set; }
        public List<selectedTypes> selectedTypes { get; set; }
        public List<selectedFootnotes> selectedFootnotes { get; set; }
    }
}
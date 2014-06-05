using System;
using System.Data.Entity;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using IndInv.Helpers;

namespace IndInv.Models
{

    public class Indicators
    {

        [Key]
        public string Indicator_ID { get; set; }
        public Int16 Area_ID { get; set; }
        [Display(Name = "Objectives/Metrics")]
        public string Indicator { get; set; }
        [Display(Name = "Indicator Type")]
        public string Indicator_Type { get; set; }
        [Display(Name = "FY 10/11")]
        public string FY_10_11 { get; set; }
        public string FY_10_11_Sup { get; set; }
        [Display(Name = "FY 11/12")]
        public string FY_11_12 { get; set; }
        public string FY_11_12_Sup { get; set; }
        [Display(Name = "FY 12/13")]
        public string FY_12_13 { get; set; }
        public string FY_12_13_Sup { get; set; }
        [Display(Name = "Q1")]
        public string FY_13_14_Q1 { get; set; }
        public string FY_13_14_Q1_Sup { get; set; }
        [Display(Name = "Q2")]
        public string FY_13_14_Q2 { get; set; }
        public string FY_13_14_Q2_Sup { get; set; }
        [Display(Name = "Q3")]
        public string FY_13_14_Q3 { get; set; }
        public string FY_13_14_Q3_Sup { get; set; }
        [Display(Name = "Q4")]
        public string FY_13_14_Q4 { get; set; }
        public string FY_13_14_Q4_Sup { get; set; }
        [Display(Name = "YTD")]
        public string FY_13_14_YTD { get; set; }
        public string FY_13_14_YTD_Sup { get; set; }
        [Display(Name = "Target")]
        public string Target { get; set; }
        public string Target_Sup { get; set; }
        [Display(Name = "Comparator")]
        public string Comparator { get; set; }
        public string Comparator_Sup { get; set; }
        [Display(Name = "Performance Threshold")]
        public string Performance_Threshold { get; set; }
        public string Performance_Threshold_Sup { get; set; }

        public Int16 Colour_ID { get; set; }
        public string Custom_YTD { get; set; }
        public string Custom_Q1 { get; set; }
        public string Custom_Q2 { get; set; }
        public string Custom_Q3 { get; set; }
        public string Custom_Q4 { get; set; }

        [Display(Name = "Definition/Calculation/Notes")]
        public string Definition_Calculation { get; set; }
        [Display(Name = "Target Rationale")]
        public string Target_Rationale { get; set; }
        [Display(Name = "Comparator Source")]
        public string Comparator_Source { get; set; }

        public string Data_Source_MSH { get; set; }
        public string Data_Source_Benchmark { get; set; }
        public string OPEO_Lead { get; set; }

        public virtual string Q1_Color { get { return Colour.getColour(FY_13_14_Q1, Target, FY_13_14_Q1_Sup, Custom_Q1, false, this); } }
        public virtual string Q2_Color { get { return Colour.getColour(FY_13_14_Q2, Target, FY_13_14_Q2_Sup, Custom_Q2, false, this); } }
        public virtual string Q3_Color { get { return Colour.getColour(FY_13_14_Q3, Target, FY_13_14_Q3_Sup, Custom_Q3, false, this); } }
        public virtual string Q4_Color { get { return Colour.getColour(FY_13_14_Q4, Target, FY_13_14_Q4_Sup, Custom_Q4, false, this); } }
        public virtual string YTD_Color { get { return Colour.getColour(FY_13_14_YTD, Target, FY_13_14_YTD, Custom_YTD, true, this); } }

        public virtual ICollection<Indicator_CoE_Maps> Indicator_CoE_Map { get; set; }
        public virtual Areas Area { get; set; }
        public virtual ICollection<Indicator_Footnote_Maps> Indicator_Footnote_Map { get; set; }
    }

    public class CoEs
    {
        [Key]
        public Int16 CoE_ID { get; set; }
        public string CoE { get; set; }
        public string CoE_Abbr { get; set; }

        public virtual ICollection<Indicator_CoE_Maps> Indicator_CoE_Map { get; set; }
        public virtual ICollection<Area_CoE_Maps> Area_CoE_Map { get; set; }
    }

    public class Indicator_CoE_Maps
    {
        [Key]
        public Int16 Map_ID { get; set; }
        public Int16 CoE_ID { get; set; }
        public string Indicator_ID { get; set; }
        [Display(Name = "#")]
        public Int16 Number { get; set; }

        public virtual CoEs CoE { get; set; }
        public virtual Indicators Indicator { get; set; }
    }

    public class Areas
    {
        [Key]
        public Int16 Area_ID { get; set; }
        public string Area { get; set; }
        public Int16 Sort { get; set; }

        public virtual ICollection<Indicators> Indicator { get; set; }
        public virtual ICollection<Area_CoE_Maps> Area_CoE_Map { get; set; }
    }

    public class Area_CoE_Maps
    {
        [Key]
        public Int16 Map_ID { get; set; }
        public Int16 CoE_ID { get; set; }
        public Int16 Area_ID { get; set; }
        public string Objective { get; set; }

        public virtual CoEs CoE { get; set; }
        public virtual Areas Area { get; set; }
    }

    public class Footnotes
    {
        [Key]
        public string Footnote_ID { get; set; }
        public string Footnote { get; set; }
        public string Footnote_Symbol { get; set; }
        public virtual ICollection<Indicator_Footnote_Maps> Indicator_Footnote_Map { get; set; }
    }

    public class Indicator_Footnote_Maps
    {
        [Key]
        public Int16 Map_ID { get; set; }
        public string Footnote_ID { get; set; }
        public string Indicator_ID { get; set; }

        public virtual Footnotes Footnote { get; set; }
        public virtual Indicators Indicator { get; set; }
    }

    public class InventoryDBContext : DbContext
    {
        public DbSet<Indicators> Indicators { get; set; }
        public DbSet<CoEs> CoEs { get; set; }
        public DbSet<Areas> Areas { get; set; }
        public DbSet<Footnotes> Footnotes { get; set; }

        public DbSet<Indicator_CoE_Maps> Indicator_CoE_Maps { get; set; }
        public DbSet<Area_CoE_Maps> Area_CoE_Maps { get; set; }
        public DbSet<Indicator_Footnote_Maps> Indicator_Footnote_Maps { get; set; }

        //public InventoryDBContext()
        //    : base("Indicator_Inventory")
        //{
        //}

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //modelBuilder.Entity<CoEs>().HasMany(c => c.Indicator_CoE_Map).WithRequired(p => p.CoE).HasForeignKey(c => c.CoE_ID);
            //modelBuilder.Entity<CoEs>().
            //  HasMany(c => c.Indicator_CoE_Map).
            //  WithMany(p => p.CoE).
            //  Map(
            //   m =>
            //   {
            //       m.MapLeftKey("CoE_ID");
            //       m.MapRightKey("Map_ID");
            //       m.ToTable("Indicator_COE_Maps");
            //   });

            //modelBuilder.Entity<Indicators>().
            //  HasMany(c => c.Indicator_CoE_Map).
            //  WithMany(p => p.Indicator).
            //  Map(
            //   m =>
            //   {
            //       m.MapLeftKey("CoE_ID");
            //       m.MapRightKey("Map_ID");
            //       m.ToTable("Indicator_COE_Maps");
            //   });

            //modelBuilder.Entity<Areas>().
            //      HasMany(c => c.CoE).
            //      WithMany(p => p.Area).
            //      Map(
            //       m =>
            //       {
            //           m.MapLeftKey("Area_ID");
            //           m.MapRightKey("CoE_ID");
            //           m.ToTable("Objective_Maps");
            //       });

        }
    }
}


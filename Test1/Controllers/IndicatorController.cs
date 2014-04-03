using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IndInv.Models;
using IndInv.Models.ViewModels;

namespace IndInv.Controllers
{

    public class IndicatorController : Controller
    {
        private InventoryDBContext db = new InventoryDBContext();

        //
        // GET: /Indicator/

        [HttpGet]
        public ActionResult Index()
        {
            var viewModel = new indexViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.ToList(),
                allAreas = db.Areas.ToList(),
                allFootnotes=db.Indicator_Footnote_Maps.ToList()
            };

            return View(viewModel);
        }

        [HttpPost]
        public ActionResult searchAdvanced(IList<searchViewModel> advancedSearch)
        {
                TempData["search"] = advancedSearch.FirstOrDefault();
                return Json(Url.Action("searchResults", "Indicator")); 
        }

        public ActionResult searchResults()
        {
            TempData.Keep();
            searchViewModel advancedSearch = (searchViewModel)TempData["search"];

            List<Indicators> indicatorList = db.Indicators.ToList();
            List<Indicators> indicatorListString = new List<Indicators>();
            string searchString = advancedSearch.searchString;
            if (searchString != null)
            {
                string[] searchStrings;
                searchStrings = searchString.Split(' ');
                foreach (var sS in searchStrings)
                {
                    indicatorList = indicatorList.Where(s => s.Indicator.ToString().ToLower().Contains(sS.ToLower())).ToList();
                    //indicatorListString.AddRange(db.Indicators.Where(s => s.Indicator.ToLower().Contains(sS.ToLower())).ToList());
                    //indicatorListString = indicatorList.Where(s => s.Indicator.ToString().ToLower().Contains(sS.ToLower())).ToList();
                    //indicatorList = indicatorListString;
                }
                //indicatorList = indicatorList.Intersect(indicatorListString).ToList();
            }

            List<Indicators> indicatorListCoE = new List<Indicators>();
            List<CoE_IDsViewModel> searchCoEs;
            searchCoEs = advancedSearch.selectedCoEs;
            if (searchCoEs != null)
            {
                foreach (var coe in searchCoEs)
                {
                    indicatorListCoE.AddRange(db.Indicators.Where(s => s.Indicator_CoE_Map.Any(x => x.CoE_ID == coe.CoE_ID)).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListCoE).ToList();
            }
            
            List<Indicators> indicatorListAreas = new List<Indicators>();
            List<Area_IDsViewModel> searchAreas;
            searchAreas = advancedSearch.selectedAreas;
            if (searchAreas != null)
            {
                foreach (var area in searchAreas)
                {
                    indicatorListAreas.AddRange(db.Indicators.Where(s => s.Area_ID == area.Area_ID).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListAreas).ToList();
            }

            List<Indicators> indicatorListTypes = new List<Indicators>();
            List<Indicator_TypesViewModel> searchTypes;
            searchTypes = advancedSearch.selectedTypes;
            if (searchTypes != null)
            {
                foreach (var type in searchTypes)
                {
                    indicatorListTypes.AddRange(db.Indicators.Where(s => s.Indicator_Type.Replace("/","").Replace("&","").Replace(" ","") == type.Indicator_Type).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListTypes).ToList();
            }

            if (ModelState.IsValid)
            {
                var viewModel = new indexViewModel
                {
                    allIndicators = indicatorList.Distinct().ToList(),
                    
                    allCoEs = db.CoEs.ToList(),
                    allAreas = db.Areas.ToList(),
                    allFootnotes = db.Indicator_Footnote_Maps.ToList()
                };
                return View(viewModel);
            }

            return View();
        }

        public ActionResult viewPR()
        {
            ModelState.Clear();
            var viewModel = new PRViewModel
            {
                //allCoEs = db.CoEs.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.ToList(),
                allFootnoteMaps = db.Indicator_Footnote_Maps.ToList()
            };

            //Response.AddHeader("Content-Disposition", "filename=thefilename.xls");
            //Response.ContentType = "application/vnd.ms-excel";

            return View(viewModel);
        }

        public ActionResult editCoEMaps()
        {
            var viewModel = new Indicator_CoE_MapsViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.ToList()
            };
            return View(viewModel);
        }

        public ActionResult editFootnoteMaps()
        {
            List<Indicator_Footnote_Maps> footnoteMaps = new List<Indicator_Footnote_Maps>();
            foreach (var footnote in db.Indicator_Footnote_Maps.OrderBy(e => e.Map_ID).ToList())
            {
                footnoteMaps.Add(footnote);
            }

            var viewModel = db.Indicators.Select(x => new Indicator_Footnote_MapsViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Indicator= x.Indicator,
            }).ToList();

            foreach (var Indicator in viewModel)
            {
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 0)
                {
                    Indicator.Footnote_ID_1 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).First().Footnote_ID;
                    Indicator.Map_ID_1 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 1)
                {
                    Indicator.Footnote_ID_2 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(1).First().Footnote_ID;
                    Indicator.Map_ID_2 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(1).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 2)
                {
                    Indicator.Footnote_ID_3 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(2).First().Footnote_ID;
                    Indicator.Map_ID_3 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(2).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 3)
                {
                    Indicator.Footnote_ID_4 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(3).First().Footnote_ID;
                    Indicator.Map_ID_4 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(3).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 4)
                {
                    Indicator.Footnote_ID_5 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(4).First().Footnote_ID;
                    Indicator.Map_ID_5 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(4).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 5)
                {
                    Indicator.Footnote_ID_6 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(5).First().Footnote_ID;
                    Indicator.Map_ID_6 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(5).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 6)
                {
                    Indicator.Footnote_ID_7 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(6).First().Footnote_ID;
                    Indicator.Map_ID_7 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(6).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 7)
                {
                    Indicator.Footnote_ID_8 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(7).First().Footnote_ID;
                    Indicator.Map_ID_8 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(7).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 8)
                {
                    Indicator.Footnote_ID_9 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(8).First().Footnote_ID;
                    Indicator.Map_ID_9 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(8).First().Map_ID;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 9)
                {
                    Indicator.Footnote_ID_10 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(9).First().Footnote_ID;
                    Indicator.Map_ID_10 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(9).First().Map_ID;
                }
            }

            return View(viewModel);
        }

        [HttpPost]
        public ActionResult editFootnoteMaps(IList<Indicator_Footnote_Maps> newMaps)
        {
            var mapID = newMaps[0].Map_ID;
            var footnoteID = newMaps[0].Footnote_ID;
            if (footnoteID == null)
            {
                var deleteMap = db.Indicator_Footnote_Maps.Find(newMaps[0].Map_ID);
                if (deleteMap != null)
                {
                    db.Indicator_Footnote_Maps.Remove(deleteMap);
                    db.SaveChanges();
                }
                return View();
            }
            else if (db.Indicator_Footnote_Maps.Any(x => x.Map_ID == mapID))
            {
                if (ModelState.IsValid && db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    db.Entry(newMaps[0]).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                else
                {
                    var oldMap = db.Indicator_Footnote_Maps.Find(newMaps[0].Map_ID);
                    var viewModel = new
                    {
                        Map_ID = oldMap.Map_ID,
                        Footnote_ID = oldMap.Footnote_ID,
                        Indicator_ID = oldMap.Indicator_ID,
                        State = "InvalidChange"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                if (ModelState.IsValid && db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    db.Indicator_Footnote_Maps.Add(newMaps[0]);
                    db.SaveChanges();
                    var viewModel = new  
                    {
                        Map_ID = newMaps[0].Map_ID,
                        Footnote_ID  = newMaps[0].Footnote_ID ,
                        Indicator_ID = newMaps[0].Indicator_ID,
                        State = "NewID"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
                else if (ModelState.IsValid && !db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    var viewModel = new
                    {
                        State = "InvalidAdd"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
            }
            return View();
        }

        [HttpGet]
        public ActionResult getIndicatorList()
        {
            var viewModel = db.Indicators.Select(x => new IndicatorListViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Indicator = x.Indicator,
            }).ToList();
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult editInventory(String Indicator_ID_Filter)
        {
            var viewModelItems = db.Indicators.ToArray();
            var viewModel = viewModelItems.OrderBy(x => x.Indicator_ID).Select(x => new InventoryViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Area_ID = x.Area_ID,
                Indicator = x.Indicator,
                FY_10_11 = x.FY_10_11,
                FY_10_11_Sup = x.FY_10_11_Sup,
                FY_11_12 = x.FY_11_12,
                FY_11_12_Sup = x.FY_11_12_Sup,
                FY_12_13 = x.FY_12_13,
                FY_12_13_Sup = x.FY_12_13_Sup,
                FY_13_14_Q1 = x.FY_13_14_Q1,
                FY_13_14_Q1_Sup = x.FY_13_14_Q1_Sup,
                FY_13_14_Q2 = x.FY_13_14_Q2,
                FY_13_14_Q2_Sup = x.FY_13_14_Q2_Sup,
                FY_13_14_Q3 = x.FY_13_14_Q3,
                FY_13_14_Q3_Sup = x.FY_13_14_Q3_Sup,
                FY_13_14_Q4 = x.FY_13_14_Q4,
                FY_13_14_Q4_Sup = x.FY_13_14_Q4_Sup,
                FY_13_14_YTD = x.FY_13_14_YTD,
                FY_13_14_YTD_Sup = x.FY_13_14_YTD_Sup,
                Target = x.Target,
                Target_Sup = x.Target_Sup,
                Comparator = x.Comparator,
                Comparator_Sup = x.Comparator_Sup,
                Performance_Threshold = x.Performance_Threshold,
                Performance_Threshold_Sup = x.Performance_Threshold_Sup,

                Colour_ID = x.Colour_ID,
                Custom_YTD = x.Custom_YTD,
                Custom_Q1 = x.Custom_Q1,
                Custom_Q2 = x.Custom_Q2,
                Custom_Q3 = x.Custom_Q3,
                Custom_Q4 = x.Custom_Q4,

                Definition_Calculation = x.Definition_Calculation,
                Target_Rationale = x.Target_Rationale,
                Comparator_Source = x.Comparator_Source,

                Data_Source_MSH = x.Data_Source_MSH,
                Data_Source_Benchmark = x.Data_Source_Benchmark,
                OPEO_Lead = x.OPEO_Lead,

                Q1_Color = x.Q1_Color,
                Q2_Color = x.Q2_Color,
                Q3_Color = x.Q3_Color,
                Q4_Color = x.Q4_Color,
                YTD_Color = x.YTD_Color
            }).ToList();
            if (Request.IsAjaxRequest())
            {
                return Json(viewModel.Where(x => x.Indicator_ID.ToString().Contains(Indicator_ID_Filter == null ? "" : Indicator_ID_Filter)), JsonRequestBehavior.AllowGet);
            }
            else
            {
                return View(viewModel);
            }
            
        }

        [HttpPost]
        public ActionResult editInventory(IList<Indicators> indicatorChange)
        {
            var indicatorID = indicatorChange[0].Indicator_ID;
            if (db.Indicators.Any(x => x.Indicator_ID == indicatorID ))
            {
                if (ModelState.IsValid)
                {
                    db.Entry(indicatorChange[0]).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                return View();
            } 
            else
            {
                if (ModelState.IsValid)
                {
                    db.Indicators.Add(indicatorChange[0]);
                    db.SaveChanges();
                    return View();
                }
                return View();
            }

        }

        //
        // GET: /Indicator/Details/5

        public ActionResult Details(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // GET: /Indicator/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Indicator/Create

        [HttpPost]
        public ActionResult Create(Indicators indicators)
        {
            if (ModelState.IsValid)
            {
                db.Indicators.Add(indicators);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(indicators);
        }

        //
        // GET: /Indicator/Edit/5

        public ActionResult Edit(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // POST: /Indicator/Edit/5

        [HttpPost, ValidateInput(false)]
        public ActionResult Edit(Indicators indicators)
        {
            if (ModelState.IsValid)
            {
                db.Entry(indicators).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(indicators);
        }

        //
        // GET: /Indicator/Delete/5

        public ActionResult Delete(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // POST: /Indicator/Delete/5

        [HttpPost, ActionName("Delete")]
        public ActionResult DeleteConfirmed(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            db.Indicators.Remove(indicators);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}